VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGesConduc 
   Caption         =   "Choferes"
   ClientHeight    =   7635
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10515
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
      Left            =   7590
      TabIndex        =   42
      Top             =   210
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3900
      TabIndex        =   40
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   41
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
      TabIndex        =   38
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   39
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5400
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
   Begin VB.Frame Frame3 
      Caption         =   "Veh�culos"
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
      Height          =   2715
      Left            =   240
      TabIndex        =   32
      Top             =   4200
      Width           =   10115
      Begin VB.Frame FrameToolAux 
         Height          =   555
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1605
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   0
            Left            =   210
            TabIndex        =   44
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
      Begin VB.CommandButton cmdAux 
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   1110
         TabIndex        =   28
         ToolTipText     =   "Buscar chofer"
         Top             =   1680
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux1 
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
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   37
         Tag             =   "Observaciones|T|S|||scoche_historia|observac|||"
         Text            =   "Observaciones"
         Top             =   1650
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.TextBox txtAux1 
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
         Left            =   2460
         MaxLength       =   10
         TabIndex        =   36
         Tag             =   "Fecha Fin|F|S|||scoche_historia|fechafin|dd/mm/yyyy||"
         Text            =   "Fecha fin"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux1 
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
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   35
         Tag             =   "Fecha Inicio|F|S|||scoche_historia|fechaini|dd/mm/yyyy||"
         Text            =   "Fecha INI"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtAux1 
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
         Left            =   240
         MaxLength       =   5
         TabIndex        =   34
         Tag             =   "Vehiculo|N|N|||scoche_historia|numeruve|0000||"
         Text            =   "Vehiculo"
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1725
         Left            =   240
         TabIndex        =   33
         Top             =   870
         Width           =   9535
         _ExtentX        =   16828
         _ExtentY        =   3043
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
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5400
      Top             =   6960
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
      Left            =   7890
      TabIndex        =   12
      Top             =   7080
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
      Left            =   9210
      TabIndex        =   13
      Top             =   7080
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
      Left            =   9210
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   240
      TabIndex        =   14
      Top             =   780
      Width           =   10095
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
         Height          =   1095
         Index           =   11
         Left            =   6330
         TabIndex        =   11
         Tag             =   "Observaciones|T|S|||schofe|observac|||"
         Text            =   "Text"
         Top             =   2130
         Width           =   3645
      End
      Begin VB.TextBox Text2 
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
         Left            =   2190
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   2760
         Width           =   3105
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
         Index           =   10
         Left            =   1290
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "Codigo situacion|N|N|||schofe|codsitua|||"
         Text            =   "Text"
         Top             =   2760
         Width           =   855
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
         Left            =   6330
         MaxLength       =   40
         TabIndex        =   7
         Tag             =   "Mail|T|S|||schofe|maichofe|||"
         Text            =   "Text"
         Top             =   1320
         Width           =   3645
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
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Movil|T|S|||schofe|movchofe|||"
         Text            =   "Text"
         Top             =   840
         Width           =   1785
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
         Left            =   6330
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Telefono|T|S|||schofe|telchofe|||"
         Text            =   "Text"
         Top             =   840
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
         Index           =   6
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   9
         Tag             =   "CIF|T|S|||schofe|cifchofe|||"
         Text            =   "Text"
         Top             =   2280
         Width           =   1485
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
         Left            =   1290
         MaxLength       =   35
         TabIndex        =   8
         Tag             =   "Provincia|T|S|||schofe|prochofe|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1800
         Width           =   4035
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
         Left            =   3360
         MaxLength       =   35
         TabIndex        =   6
         Tag             =   "Poblaci�n|T|S|||schofe|pobchofe|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1320
         Width           =   1965
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
         Left            =   1290
         MaxLength       =   5
         TabIndex        =   5
         Tag             =   "CP|T|S|||schofe|codpobla|||"
         Text            =   "Text"
         Top             =   1320
         Width           =   975
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
         Left            =   1290
         MaxLength       =   35
         TabIndex        =   2
         Tag             =   "Domicilio|T|S|||schofe|domchofe|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   840
         Width           =   4035
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
         Left            =   4080
         MaxLength       =   35
         TabIndex        =   1
         Tag             =   "Nombre chofer|T|N|||schofe|nomchofe|||"
         Text            =   "Text"
         Top             =   360
         Width           =   5895
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
         Index           =   0
         Left            =   1290
         MaxLength       =   5
         TabIndex        =   0
         Tag             =   "Codigo chofer|N|N|||schofe|codchofe||S|"
         Text            =   "Text"
         Top             =   360
         Width           =   870
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   1
         Left            =   6060
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   6960
         Tag             =   "-1"
         ToolTipText     =   "Ver observaciones"
         Top             =   1830
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1050
         Tag             =   "-1"
         ToolTipText     =   "Buscar situaci�n"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Situaci�n:"
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
         TabIndex        =   30
         Top             =   2760
         Width           =   915
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
         Index           =   11
         Left            =   5400
         TabIndex        =   29
         Top             =   1320
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Movil"
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
         Left            =   7680
         TabIndex        =   27
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tel�fono:"
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
         Left            =   5400
         TabIndex        =   26
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "CIF:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia:"
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
         TabIndex        =   24
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Poblaci�n:"
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
         Left            =   2370
         TabIndex        =   23
         Top             =   1320
         Width           =   945
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1020
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "CP:"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio:"
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
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
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
         Left            =   5370
         TabIndex        =   17
         Top             =   1830
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
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
         Left            =   3240
         TabIndex        =   16
         Top             =   390
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   240
      TabIndex        =   15
      Top             =   6960
      Width           =   3165
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   210
         Width           =   2745
      End
   End
   Begin VB.Menu mnopciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
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
      Begin VB.Menu mnnuevo 
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
      Begin VB.Menu mnsalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmGesConduc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 403

Public WithEvents frmB As frmBasico2
Attribute frmB.VB_VarHelpID = -1
Public WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Public WithEvents frmS As frmFacSituaciones
Attribute frmS.VB_VarHelpID = -1
Public WithEvents frmV As frmGesVSocio
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

Private Sub cmdAceptar_Click()
Dim I As Integer
Dim CadB As String
Dim cad As String

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
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If InsertarLinea Then
                    CargaGrid DataGrid1, Adodc2
                    BotonAnyadirLinea
                End If
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    CargaTxtAux False, False
                    CargaGrid DataGrid1, Adodc2
                    ModificaLineas = 0
'                    PonerBotonCabecera True
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next

    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    
    If b Then
        Me.lblIndicador(0).Caption = "L�neas "
        PonerFocoBtn Me.cmdRegresar
    End If
    
    'Habilitar las opciones correctas del menu seg�n Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu seg�n Nivel de Acceso
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Pedido: sliped
Dim Sql As String

On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    
    conn.BeginTrans
    
    If DatosOkLinea() Then
        'Creamos la sentencia SQL
        Sql = "UPDATE schofe_historia Set numeruve = " & txtAux1(0).Text & ", fechaini='" & Format(txtAux1(1).Text, FormatoFecha) & "', "
        Sql = Sql & "fechafin='" & Format(txtAux1(2).Text, FormatoFecha) & "', observac=" & DBSet(txtAux1(3).Text, "T")
        Sql = Sql & " where codchofe=" & Adodc2.Recordset!codchofe & " AND numlinea=" & Adodc2.Recordset!numlinea
        
        conn.Execute Sql
        
        ' modificamos la tabla de choferes del socio
        Sql = "update sclien_chofer, sclien set sclien_chofer.fechaalt = " & DBSet(txtAux1(1).Text, "F") & ", sclien_chofer.fechabaj = " & DBSet(txtAux1(2).Text, "F", "S")
        Sql = Sql & ", sclien_chofer.obsevac = " & DBSet(txtAux1(3).Text, "T") & " where codchofe = " & Me.Adodc1.Recordset!codchofe
        Sql = Sql & " and sclien.numeruve = " & DBSet(txtAux1(0).Text, "N")
        Sql = Sql & " and sclien.codclien = sclien_chofer.codsocio "
        
        conn.Execute Sql
        
        ModificarLinea = True
    End If
    
    conn.CommitTrans
    Exit Function

EModificarLinea:
    conn.RollbackTrans
    MuestraError Err.Number, "Modificar Lineas Pedido" & vbCrLf & Err.Description
End Function

Private Function DatosOkLinea() As Boolean
Dim Sql As String
Dim Socio As String

    DatosOkLinea = False
    If txtAux1(0).Text <> "" Then
        If Not IsNumeric(txtAux1(0).Text) Then
            DatosOkLinea = False
            Exit Function
        End If
        
        ' si queremos insertar
        If ModificaLineas = 1 Then
            ' solo si no tiene fecha de baja hacemos las comprobaciones
            If txtAux1(2).Text = "" Then
                ' comprobamos no sea conductor una V sin fecha de baja
                Sql = "select count(*) from schofe_historia where codchofe = " & Me.Adodc1.Recordset!codchofe
                'SQL = SQL & " and numeruve = " & DBSet(txtAux1(0).Text, "N")
                Sql = Sql & " and (fechafin is null or fechafin = '0000-00-00') "
                
                If TotalRegistros(Sql) > 0 Then
                    MsgBox "Este chofer ya est� asociado a una V sin fecha de baja. Revise.", vbExclamation
                    Exit Function
                Else
                    ' comprobamos si la v est� asignada a algun socio o no
                    Socio = DevuelveValor("select codclien from sclien where numeruve= " & DBSet(txtAux1(0).Text, "N"))
                    If Socio = 0 Then
                        MsgBox "Esta V no est� asignada en este momento a ning�n socio. Revise.", vbExclamation
                        Exit Function
                    End If
                End If
            End If
        End If
    Else
        DatosOkLinea = False
        MsgBox "Es necesario introducir un codigo de veh�culo.", vbExclamation
        Exit Function
    End If

    If txtAux1(1).Text = "" Then
        DatosOkLinea = False
        MsgBox "Es necesario introducir la fecha de inicio. Revise", vbExclamation
        Exit Function
    End If

    DatosOkLinea = True

End Function

Private Function InsertarLinea() As Boolean
Dim Sql As String
Dim vWhere As String
Dim Socio As Long
Dim numF As String

On Error GoTo EInsertarLinea

    conn.BeginTrans


    InsertarLinea = False
    Sql = ""
    If DatosOkLinea Then
        
        vWhere = "codchofe=" & Val(Text1(0).Text)
        numF = SugerirCodigoSiguienteStr("schofe_historia", "numlinea", vWhere)
        
        Sql = "INSERT INTO schofe_historia "
        Sql = Sql & "(codchofe, numlinea, numeruve, fechaini,fechafin,observac) "
        Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
        Sql = Sql & DBSet(txtAux1(0).Text, "T") & ",'" & Format(txtAux1(1).Text, FormatoFecha) & "','"
        Sql = Sql & Format(txtAux1(2).Text, FormatoFecha) & "'," & DBSet(txtAux1(3).Text, "T") & ")"
        
        conn.Execute Sql
        
        ' insertamos en la tabla de choferes del socio
        Socio = DevuelveValor("select codclien from sclien where numeruve= " & DBSet(txtAux1(0).Text, "N"))
        
        vWhere = "codsocio=" & DBSet(Socio, "N")
        numF = SugerirCodigoSiguienteStr("sclien_chofer", "numlinea", vWhere)
        
        Sql = "INSERT INTO sclien_chofer "
        Sql = Sql & "(codsocio,numlinea,codchofe,fechaalt,fechabaj,obsevac) "
        Sql = Sql & "VALUES (" & DBSet(Socio, "N") & ", " & numF & ","
        Sql = Sql & DBSet(Text1(0).Text, "T") & "," & DBSet(txtAux1(1).Text, "F") & ","
        Sql = Sql & DBSet(txtAux1(2).Text, "F", "S") & "," & DBSet(txtAux1(3).Text, "T") & ")"
        
        conn.Execute Sql
        
        InsertarLinea = True
    End If
    
    conn.CommitTrans
    Exit Function
    
EInsertarLinea:
    conn.RollbackTrans
    MuestraError Err.Number, "Insertar Lineas Trabajador" & vbCrLf & Err.Description
    
End Function

Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    cadB1 = ObtenerBusqueda(Me, True)
    
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
Private Function DatosOk() As Boolean

DatosOk = False

'CODIGO DE CHOFER
If Text1(0).Text = "" Then
    MsgBox "Debe introducir el c�digo de chofer.", vbExclamation
    PonerFoco Text1(0)
    Exit Function
ElseIf Not IsNumeric(Text1(0).Text) Then
        Exit Function
End If

'NOMBRE DEL CHOFER
If Text1(1).Text = "" Then
    MsgBox "Debe introducir el nombre del chofer.", vbExclamation
    PonerFoco Text1(1)
    Exit Function
End If

'CODIGO SITUACI�N
If Text1(10).Text = "" Then
    MsgBox "Debe introducir el c�digo de situaci�n.", vbExclamation
    PonerFoco Text1(10)
    Exit Function
ElseIf Not IsNumeric(Text1(10).Text) Then
        Exit Function
End If

DatosOk = True
    
End Function

Private Sub cmdAux_Click(Index As Integer)
    Select Case Index
        Case 0
            Set frmV = New frmGesVSocio
            frmV.DatosADevolverBusqueda = "0|"
            frmV.Show vbModal
            Set frmV = Nothing
    End Select
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador(0).Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
        If DataGrid1.Row >= 0 Then
            DeseleccionaGrid DataGrid1
            DataGrid1.Bookmark = 1
        End If
        cmdRegresar.Caption = "Regresar"
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Adodc1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Adodc1.Recordset.Fields(0) & "|"
        cad = cad & Adodc1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del form
    Me.Icon = frmppal.Icon
    

    'ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 13
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Bot�n Buscar
'        .Buttons(2).Image = 2   'Bot�n Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        .Buttons(9).Image = 10 'Lineas
'        .Buttons(10).Image = 16 'Imprmir
'        .Buttons(11).Image = 15 'Salir
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
    
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next kCampo
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
        
    ImgMail(1).Picture = frmppal.imgIcoForms.ListImages(4).Picture
    LimpiarDataGrids
    
    '## A mano
    NombreTabla = "schofe"
    Ordenacion = " ORDER BY codchofe"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = "Select * from " & NombreTabla & " where codchofe=-1"
    Adodc1.Refresh
    
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
    
End Sub

Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo


    'Actualiza Iconos Insertar,Modificar,Eliminar
'    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    
    Modo = Kmodo
    PonerIndicador lblIndicador(0), Modo, ModificaLineas
    
    
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
    If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Me.Adodc1.Recordset.RecordCount > 1

    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    Me.cmdAux(0).Enabled = (Modo = 5 And ModificaLineas = 1)
    
    '-----------------------------
    ' ***************************
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
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
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
        
        
        'lineas
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


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean
Dim I As Byte
Dim bAux As Boolean

    b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    
''   boton de lineas
'    Toolbar1.Buttons(9).Enabled = (Modo = 2)
    '------------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b

    b = (Modo = 2)
    For I = 0 To ToolAux.Count - 1 '[Monica]30/09/2013: antes - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adodc2.Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador(0).Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    
    LimpiarDataGrids
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub frmB_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 2)
        CadB = CadB & " AND " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
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

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
'Situaciones
    Text1(10).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo Text1(10)
    Text2.Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmV_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

Select Case Index
    Case 2
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(11).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Adodc1.Recordset!observac, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(11).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
    Case 0
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                indice = 4
            Else
                PonerFoco Text1(3)
            End If
    Case 1
            indice = 10
            Set frmS = New frmFacSituaciones
            frmS.DatosADevolverBusqueda = "0"
             If Not IsNumeric(Text1(indice)) Then Text1(indice).Text = ""
            frmS.Show vbModal
            Set frmS = Nothing
    End Select
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

If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
If Index <> 9 And Index <> 11 Then Text1(Index) = UCase(Text1(Index).Text)

Select Case Index
    Case 10
        If Text1(Index).Text <> "" Then
            If IsNumeric(Text1(Index).Text) Then
                encontrado = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", Text1(Index).Text, "T")
                If encontrado <> "" Then
                    Text2.Text = encontrado
                Else
                    MsgBox "El c�digo de situaci�n introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                MsgBox "El c�digo de situaci�n debe ser num�rico.", vbExclamation
                PonerFoco Text1(Index)
            End If
        End If
    Case 6 'nif
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
        End If
End Select
End Sub

Private Sub LimpiarDataGrids()
Dim Sql As String
'Pone los Grids sin datos, apuntando a ning�n registro
On Error Resume Next

    Sql = "select * from schofe_historia where codchofe=-1"
    CargaGridGnral DataGrid1, Adodc2, Sql, PrimeraVez
    CargaGrid DataGrid1, Adodc2
    
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
            If ModificaLineas = 1 Then 'INSERTAR
                ModificaLineas = 0
                DataGrid1.AllowAddNew = False
                If Not Adodc2.Recordset.EOF Then Adodc2.Recordset.MoveFirst
            Else
                ModificaLineas = 0
            End If
'            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
            PonerModo 2
            PonerCampos
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

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    
    ModificaLineas = 0
    
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea
        Case 2
            BotonModificarLinea
        Case 3
            BotonEliminarFila
        Case Else
    End Select
    'End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Nuevo
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Buscar
           mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 8  'imprimir
            printNou
    End Select
End Sub
Private Sub mnLineas_Click()
    BotonMtoLineas "Hist�rico"
End Sub

Private Sub BotonMtoLineas(cad As String)
    ModificaLineas = 0
    PonerModo 5
    PonerBotonCabecera True
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Cabecera de Pedidos
         BotonAnyadir
    End If
End Sub

Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    
    PonerModo 5
'    PonerBotonCabecera False
'    lblIndicador(0).Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Adodc2
    CargaTxtAux True, True
   
    PonerFoco txtAux1(0)
    Me.DataGrid1.Enabled = False

End Sub


Private Sub CargaTxtAux(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim I As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For I = 0 To txtAux1.Count - 1 'TextBox
            txtAux1(I).top = 290
            txtAux1(I).visible = visible
        Next I
        cmdAux(0).top = 290
        cmdAux(0).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For I = 0 To txtAux1.Count - 1
                txtAux1(I).Text = ""
                BloquearTxt txtAux1(I), False
            Next I
        Else 'Vamos a modificar
            For I = 0 To txtAux1.Count - 1
                txtAux1(I).Text = DataGrid1.Columns(I + 2).Text
                If I = 0 Then
                    txtAux1(I).Locked = True
                    txtAux1(I).BackColor = &H80000005
                    cmdAux(I).Enabled = False
                Else
                    txtAux1(I).Locked = False
                End If
            Next I
        End If
        

        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For I = 0 To txtAux1.Count - 1
            txtAux1(I).top = alto
            txtAux1(I).Height = 330 ' DataGrid1.RowHeight
        Next I
        
        
'        'Fijamos anchura y posicion Left
'        '--------------------------------
'        'vehiculo
'        txtAux1(0).Left = DataGrid1.Left + 330
'        txtAux1(0).Width  = DataGrid1.Columns(2).Width - 100
'        'fecha ini
'        txtAux1(1).Width = DataGrid1.Columns(3).Width - 100
'        txtAux1(1).Left = txtAux1(0).Left + (txtAux1(0).Width + 100)
'        'fecha fin
'        txtAux1(2).Width = DataGrid1.Columns(4).Width - 100
'        txtAux1(2).Left = txtAux1(1).Left + (txtAux1(1).Width + 100)
'        'observaciones
'        txtAux1(3).Width = DataGrid1.Columns(5).Width - 3200
'        txtAux1(3).Left = txtAux1(2).Left + (txtAux1(2).Width + 100)
        
        'uve
        txtAux1(0).Left = DataGrid1.Left + 330
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 160
        cmdAux(0).Left = txtAux1(0).Left + txtAux1(0).Width - 50
        
        'fecha ini
        txtAux1(1).Left = cmdAux(0).Left + cmdAux(0).Width + 40
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 40
        
        'fecha baja
        txtAux1(2).Left = txtAux1(1).Left + txtAux1(1).Width + 55
        txtAux1(2).Width = DataGrid1.Columns(4).Width - 60
        
        'observaciones
        txtAux1(3).Left = txtAux1(2).Left + txtAux1(2).Width + 35
        txtAux1(3).Width = DataGrid1.Columns(5).Width - 30
        
        'Los ponemos Visibles o No
        '--------------------------
        For I = 0 To txtAux1.Count - 1
            txtAux1(I).visible = visible
        Next I
        Me.cmdAux(0).Height = Me.DataGrid1.RowHeight
        Me.cmdAux(0).top = alto
        Me.cmdAux(0).visible = visible
    End If
End Sub
Private Sub BotonAnyadir()
'A�adir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vac�a los TextBox
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    PonerFoco Text1(0)
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
    
    If Adodc2.Recordset.EOF Then Exit Sub
    vWhere = "codchofe=" & Adodc2.Recordset!codchofe & " and numlinea=" & Adodc2.Recordset!numlinea
    
    If Not BloqueaRegistro("schofe_historia", vWhere) Then Exit Sub
    
    CargaTxtAux True, False
    ModificaLineas = 2 'Modificar
    
    PonerModo 5
    
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador(0).Caption = "MODIFICAR"
'    PonerBotonCabecera False
        
    PonerFoco txtAux1(1)
    Me.DataGrid1.Enabled = False
    
EModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
Private Sub BotonModificar()
'Prepara el Form para Modificar la cabecera de Pedidos (tabla: scaped)
Dim DeVarios As Boolean
Dim Sql As String
On Error GoTo EModificar

    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1)
           
   
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
Private Sub mnSalir_Click()
    Unload Me
End Sub
Private Sub BotonEliminar()
Dim msg As String
Dim Sql As String
Dim SQL2 As String
Dim Rs As ADODB.Recordset

On Error GoTo EEliminar

msg = "Esta seguro que desea eliminar el chofer:" & Text1(0).Text & "?"
If MsgBox(msg, vbYesNo) = vbYes Then
    NumRegElim = Adodc1.Recordset.AbsolutePosition
    
    conn.BeginTrans
    
    Sql = "select * from schofe_historia where codchofe = " & DBSet(Text1(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        SQL2 = "delete from sclien_chofer where codchofe = " & DBSet(Text1(0).Text, "N")
        SQL2 = SQL2 & " and codsocio in (select codclien from sclien where numeruve = " & DBSet(txtAux1(0).Text, "N")
        SQL2 = SQL2 & " and fechaalt = " & DBSet(txtAux1(1).Text, "F")
        If txtAux1(2).Text = "" Then
            SQL2 = SQL2 & " and (fechabaj is null or fechabaj = '0000-00-00') "
        Else
            SQL2 = SQL2 & " and fechabaj = " & DBSet(txtAux1(2).Text, "F")
        End If
        
        conn.Execute SQL2
    Wend
    Set Rs = Nothing
    
    'Ahora borramos las lineas
    Sql = "Delete from schofe_historia where codchofe=" & Text1(0).Text
    conn.Execute Sql

    ' borramos cabecera
    Sql = "Delete from schofe where codchofe=" & Text1(0).Text
    conn.Execute Sql

    conn.CommitTrans
End If

If SituarDataTrasEliminar(Adodc1, NumRegElim) Then
    PonerCampos
End If

EEliminar:
If Err.Number <> 0 Then
    conn.RollbackTrans
    MsgBox "Error al eliminar conductor." & Err.Description
End If
End Sub

Private Sub BotonEliminarFila()
Dim msg As String
Dim Sql As String

On Error GoTo EEliminarLineas

msg = "Esta seguro que desea eliminar la linea?"

If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
    conn.BeginTrans

    ' Eliminamos las lineas de choferes en la ficha de socios
    Sql = "delete from sclien_chofer where codchofe = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and codsocio in (select codsocio from sclien where numeruve = " & DBSet(Me.Adodc2.Recordset!NumerUve, "N") & ")"
    Sql = Sql & " and fechaalt = " & DBSet(Me.Adodc2.Recordset!FechaIni, "F")
    
    If Me.Adodc2.Recordset!FechaFin = "" Then
        Sql = Sql & " and (fechabaj is null or fechabaj = '0000-00-00') "
    Else
        Sql = Sql & " and fechabaj = " & DBSet(Me.Adodc2.Recordset!FechaFin, "F")
    End If
    
    conn.Execute Sql
    
    Sql = "Delete from schofe_historia where codchofe=" & Text1(0).Text & " and numeruve = " & DBSet(Me.Adodc2.Recordset!NumerUve, "N")
    Sql = Sql & " and numlinea = " & DBSet(Me.Adodc2.Recordset!numlinea, "N")
    
    conn.Execute Sql

    conn.CommitTrans
    
    PonerModo 2
    
    CargaGrid DataGrid1, Adodc2
End If

EEliminarLineas:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        MsgBox "Error al eliminar Lineas." & Err.Description
    End If
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
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
    Else
        HacerBusqueda
        If Adodc1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Adodc1.Recordset.EOF Then Exit Sub
    DesplazamientoData Adodc1, Index, True
    PonerCampos
End Sub
Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Adodc1.RecordSource = CadenaConsulta
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbLightBlue 'vbYellow
        End If
        Exit Sub
    Else
        Adodc1.Recordset.MoveFirst
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

    
    If Adodc1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Adodc1
    
    
    If Text1(10).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomsitua", "ssitua", "codsitua", Text1(10).Text, "T")
        Text2.Text = encontrado
    End If
    
    'data2 para el grid de las lineas
    Adodc2.ConnectionString = conn
    Adodc2.RecordSource = "Select * from schofe_historia where codchofe=" & Text1(0).Text
    Adodc2.Refresh
    
    CargaGrid DataGrid1, Adodc2
    
    '-- Esto permanece para saber donde estamos
    lblIndicador(0).Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim I As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    Set vDataGrid.DataSource = vData
    vDataGrid.Columns(0).visible = False 'codcoche
    vDataGrid.Columns(1).visible = False 'numlinea

    vDataGrid.Columns(2).Caption = "Veh�culo"
    vDataGrid.Columns(2).Width = 1000
    vDataGrid.Columns(2).NumberFormat = "0000"
    vDataGrid.Columns(3).Caption = "Fecha Inicio"
    vDataGrid.Columns(3).Width = 1400
    vDataGrid.Columns(4).Caption = "Fecha Fin"
    vDataGrid.Columns(4).Width = 1400
    vDataGrid.Columns(5).Caption = "Observaciones"
    vDataGrid.Columns(5).Width = 5150
    
    
    
    vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    For I = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(I).Locked = True
        vDataGrid.Columns(I).AllowSizing = False
    Next I
    
    vDataGrid.ScrollBars = dbgAutomatic
    
    PonerModoOpcionesMenu Modo
    
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Cad = Cad & ParaGrid(Text1(0), 14, "C�digo")
'    Cad = Cad & ParaGrid(Text1(1), 65, "Nombre")
'
'    Tabla = "schofe"
'    Titulo = "Conductores"
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
'        frmB.vConexionGrid = conAri
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault

    Set frmB = New frmBasico2
    
    AyudaChoferes frmB, Text1(0).Text, CadB
    
    Set frmB = Nothing

End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 1: dirMail = Text1(9).Text
    End Select

    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codchofe=" & Text1(0).Text & ")"
    If SituarData(Adodc1, cad, Indicador) Then
       PonerModo 2
       lblIndicador(0).Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFoco txtAux1(Index), Modo
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

    Select Case Index
        Case 1
            PonerFormatoFecha txtAux1(Index)
        Case 2
            PonerFormatoFecha txtAux1(Index)
        Case 0
            If txtAux1(Index).Text <> "" Then
'                encontrado = DevuelveDesdeBD(conAri, "nomchofe", "scoche", "codcoche", txtAux1(Index).Text, "T")
'                If encontrado = "" Then
'                    MsgBox "El c�digo de vehiculo introducido no existe.", vbExclamation
'                    PonerFoco txtAux1(Index)
'                End If
                If Not IsNumeric(txtAux1(Index).Text) Then
                    MsgBox "El c�digo de vehiculo debe ser num�rico.", vbExclamation
                    PonerFoco txtAux1(Index)
                End If
            End If
    End Select
End Sub

Private Sub printNou()

    With frmImprimir2
        .cadTabla2 = "schofe"
        .Informe2 = "rGesChofer.rpt"
        If cadB1 <> "" Then
            .cadRegSelec = cadB1 'SQL2SF(cadB1)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Adodc1, Me)
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


