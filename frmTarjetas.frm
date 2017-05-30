VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTarjetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tarjetas Clientes"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12810
   Icon            =   "frmTarjetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3645
      Left            =   90
      TabIndex        =   15
      Top             =   1410
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6429
      _Version        =   393216
      Tabs            =   1
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
      TabCaption(0)   =   "Tarjetas"
      TabPicture(0)   =   "frmTarjetas.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(10)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(11)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(12)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(13)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "imgFich(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DataGrid1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtAux(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtAux(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtAux(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtAux(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtAux(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtAux(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtAux(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtAux(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdAux(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdAux(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAux(7)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtAux(6)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
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
         Height          =   360
         Index           =   6
         Left            =   8940
         MaxLength       =   40
         TabIndex        =   22
         Tag             =   "Pista 1|T|S|||slitar|pistagr1|||"
         Text            =   "pistagr1"
         Top             =   900
         Width           =   3555
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
         Height          =   360
         Index           =   7
         Left            =   8940
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   23
         Tag             =   "Pista 2|T|S|||slitar|pistagr2|||"
         Text            =   "pistagr2"
         Top             =   1620
         Width           =   3555
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   4020
         TabIndex        =   27
         ToolTipText     =   "Buscar artículo"
         Top             =   2790
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   0
         Left            =   2370
         TabIndex        =   26
         ToolTipText     =   "Buscar almacen"
         Top             =   2790
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
         Height          =   330
         Index           =   2
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   18
         Tag             =   "Fecha Caducidad|F|S|||slitar|fechacad|dd/mm/yyyy||"
         Text            =   "Fecha Caducidad"
         Top             =   2790
         Visible         =   0   'False
         Width           =   1335
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
         Height          =   360
         Index           =   9
         Left            =   5010
         MaxLength       =   200
         TabIndex        =   25
         Tag             =   "Fichero|T|S|||slitar|nomfiche|||"
         Text            =   "nomfiche"
         Top             =   3060
         Width           =   7485
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
         Height          =   360
         Index           =   8
         Left            =   8970
         MaxLength       =   40
         TabIndex        =   24
         Tag             =   "Pista 3|T|S|||slitar|pistagr3|||"
         Text            =   "pistagr3"
         Top             =   2340
         Width           =   3555
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
         Height          =   360
         Index           =   5
         Left            =   5010
         MaxLength       =   40
         TabIndex        =   21
         Tag             =   "Texto 3|T|S|||slitar|textoa3|||"
         Text            =   "textoa3"
         Top             =   2340
         Width           =   3825
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
         Height          =   360
         Index           =   4
         Left            =   5010
         MaxLength       =   40
         TabIndex        =   20
         Tag             =   "Texto 2|T|S|||slitar|textoa2|||"
         Text            =   "textoa2"
         Top             =   1620
         Width           =   3825
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
         Height          =   360
         Index           =   3
         Left            =   5010
         MaxLength       =   30
         TabIndex        =   19
         Tag             =   "Texto 1|T|S|||slitar|textoa1|||"
         Text            =   "textoa1"
         Top             =   900
         Width           =   3825
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
         Index           =   1
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Fecha Emision|F|S|||slitar|fechaemi|dd/mm/yyyy||"
         Text            =   "Fecha Emision"
         Top             =   2790
         Visible         =   0   'False
         Width           =   1125
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
         Index           =   0
         Left            =   480
         MaxLength       =   16
         TabIndex        =   16
         Tag             =   "Tarjeta|T|S|||slitar|numtarje|||"
         Text            =   "Tarjeta"
         Top             =   2790
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmTarjetas.frx":0028
         Height          =   2040
         Left            =   150
         TabIndex        =   28
         Top             =   660
         Width           =   4710
         _ExtentX        =   8308
         _ExtentY        =   3598
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
      Begin VB.Image imgFich 
         Height          =   240
         Index           =   0
         Left            =   7590
         Picture         =   "frmTarjetas.frx":003D
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero de configuración"
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
         Left            =   5010
         TabIndex        =   35
         Top             =   2820
         Width           =   2505
      End
      Begin VB.Label Label1 
         Caption         =   "Pista 3"
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
         Left            =   8970
         TabIndex        =   34
         Top             =   2100
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Pista 2"
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
         Left            =   8970
         TabIndex        =   33
         Top             =   1380
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Pista 1"
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
         Left            =   8970
         TabIndex        =   32
         Top             =   660
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Texto a imprimir (3)"
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
         Left            =   5010
         TabIndex        =   31
         Top             =   2100
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Texto a imprimir (2)"
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
         Left            =   5010
         TabIndex        =   30
         Top             =   1380
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Texto a imprimir (1)"
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
         Left            =   5010
         TabIndex        =   29
         Top             =   660
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   5085
      Width           =   2325
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
         TabIndex        =   7
         Top             =   180
         Width           =   1845
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
      Left            =   11580
      TabIndex        =   10
      Top             =   5220
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
      Left            =   10380
      TabIndex        =   9
      Top             =   5220
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   300
      Top             =   5130
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
      TabIndex        =   8
      Top             =   0
      Width           =   12810
      _ExtentX        =   22595
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
            Object.ToolTipText     =   "Tarjetas Cliente"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar Tarjeta"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Importar Fichero"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Marcar facturar"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Tarjeta"
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
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6960
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   1020
      Top             =   5160
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
      Left            =   11580
      TabIndex        =   5
      Top             =   5220
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   390
      Width           =   12585
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
         Left            =   210
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Cod. Cliente|N|N|0|999999|scatar|codclien|000000|S|"
         Text            =   "Text1"
         Top             =   420
         Width           =   760
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
         Index           =   0
         Left            =   1035
         MaxLength       =   40
         TabIndex        =   4
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   435
         Width           =   4080
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
         Left            =   6330
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "Nombre Usuario|T|N|||scatar|nomusuar||N|"
         Top             =   450
         Width           =   5625
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
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
         Left            =   5250
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Usuario|N|N|||scatar|codusuar|000000|S|"
         Text            =   "Text1 7"
         Top             =   450
         Width           =   885
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1050
         ToolTipText     =   "Buscar socio"
         Top             =   180
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
         Left            =   210
         TabIndex        =   14
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
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
         Left            =   5310
         TabIndex        =   13
         Top             =   210
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
         Index           =   50
         Left            =   6330
         TabIndex        =   12
         Top             =   210
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         HelpContextID   =   2
         Shortcut        =   ^L
      End
      Begin VB.Menu mnBuscarTarjeta 
         Caption         =   "Buscar &Tarjeta"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnImportarFichero 
         Caption         =   "Importar Fichero"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmTarjetas"
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

Private WithEvents frmC As frmFacClientes 'Form Mto Clientes
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
Dim BusquedaTarjetas As Boolean


Private Sub cmdAceptar_Click()
Dim PrimeraLin As Boolean 'Si se inserta la primera linea no esta creado el datagrid1 entonces llamar
                          ' a DataGrid, sino llamar solo a DataGrid2
Dim numlinea As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            If BusquedaTarjetas Then
                HacerBusquedaTarjeta
                BusquedaTarjetas = False
            Else
                HacerBusqueda
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                InsertarCabecera
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificarCabAlbaran Then
                    TerminaBloquear
                    
                    PosicionarData
                End If
            End If
            
         Case 5 'InsertarModificar LINEA
            'Actualizar el registro en la tabla de lineas 'slialb'
            If ModificaLineas = 1 Then 'INSERTAR lineas Albaran
                PrimeraLin = False
                If Data2.Recordset.EOF = True Then PrimeraLin = True
                If InsertarLinea(numlinea, False) Then
                    'Comprobar si el Articulo tiene control de Nº de Serie
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
                    TerminaBloquear
                    NumRegElim = Val(Data2.Recordset!numlinea)
                    CargaTxtAux False, False
                    CargaGrid2 DataGrid1, Data2
                    PosicionarData2
                    ModificaLineas = 0
                    PonerBotonCabecera True
                End If
                Me.DataGrid1.Enabled = True
            End If
    End Select
    Screen.MousePointer = vbDefault
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
                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
            End If
            ModificaLineas = 0
            PonerBotonCabecera True
            Me.DataGrid1.Enabled = True
    
            If Not Data2.Recordset.EOF Then
                CargaForaGrid
            Else
                LimpiarCampos
            End If
    
    
    End Select
        
    
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
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3



    Text1(0).Text = ""
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
       
    ModificaLineas = 1 'Ponemos Modo Añadir Linea
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    AnyadirLinea DataGrid1, Data2
    CargaTxtAux True, True
    
    PonerFoco txtAux(0)
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
    
End Sub

Private Sub BotonBuscarTarjeta()
    'Buscar
    BusquedaTarjetas = True
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        CargaTxtAux True, True
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco txtAux(0)
        txtAux(0).BackColor = vbLightBlue 'vbYellow
'        Text1(30).visible = False
'        Combo1.visible = True
    Else
        HacerBusquedaTarjeta
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue 'vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
    
End Sub




Private Sub BotonVerTodos()
Dim Cad As String
'    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select * from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " " & Ordenacion
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
Dim DeVarios As Boolean

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(2)
   
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
    
    
    ModificaLineas = 2 'Modificar
    'Añadiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
    
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

    Cad = "Usuarios de Tarjetas." & vbCrLf
    Cad = Cad & "------------------------------------       " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar el Usuario:            "
    Cad = Cad & vbCrLf & "Cliente:  " & Format(Text1(0).Text, "000000") & " " & Text2(0).Text
    Cad = Cad & vbCrLf & "Usuario:  " & Format(Text1(1).Text, "000000") & " " & Text1(2).Text
    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "
      
    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        Screen.MousePointer = vbHourglass
        
        NumRegElim = Data1.Recordset.AbsolutePosition
        NumAlbElim = Data1.Recordset.Fields(1).Value
        
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
    Sql = "¿Seguro que desea eliminar la tarjeta de usuario?     "
    Sql = Sql & vbCrLf & "Tarjeta:  " & Data2.Recordset!NUMTARJE & vbCrLf
    Sql = Sql & "Cliente:  " & Format(Data2.Recordset!CodClien, "000")
    Sql = Sql & vbCrLf & "Usuario:  " & Data2.Recordset!codusuar
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = Data2.Recordset.AbsolutePosition
        If EliminarLinea Then
            LimpiarDatosTarjeta
            ModificaLineas = 0
            CargaGrid2 DataGrid1, Data2
            SituarDataTrasEliminar Data2, NumRegElim
        End If
'        CancelaADODC
    End If
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


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim RS As ADODB.Recordset
Dim Sql As String
Dim i As Integer

    On Error GoTo Error1

    If Not Data2.Recordset.EOF Then
        CargaForaGrid
        Exit Sub
    Else
        For i = 0 To txtAux.Count - 1
            txtAux(i).Text = ""
        Next i
    End If
    Exit Sub
    
Error1:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    PrimeraVez = False
    
'    If AlbAvisoGenerado > 0 Then
'        PonerCadenaBusqueda
'        'Simulo que pulsa lineas
'        mnLineas_Click
'
'        'Simulo que le da a insertar nueva
'        mnNuevo_Click
'
'        'AlbAvisoGenerado
'        AlbAvisoGenerado = 0
'    End If
        
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon


    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 19
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 10 'Mto Lineas Ofertas
        .Buttons(11).Image = 33 'Nº Serie si lineas con articulos de control Nº serie
        .Buttons(12).Image = 18 'Importar Fichero
        .Buttons(13).Image = 30 'Marcar a facturar
        
        .Buttons(14).Image = 27 'Imprimir portes
        .Buttons(15).Image = 16 'Imprimir Pedido
'
        .Buttons(16).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
        
        If vParamAplic.ArtPortes = "" Then
            .Buttons(14).Style = tbrSeparator
            .Buttons(14).ToolTipText = ""
        Else
            .Buttons(14).Style = tbrDefault
            .Buttons(14).ToolTipText = "Imprimir portes"
        End If
    End With
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    
    VieneDeBuscar = False
    CodTipoMov = hcoCodTipoM
    
        
    '## A mano
    
    NombreTabla = "scatar"
    NomTablaLineas = "slitar" 'Tabla lineas de Albaranes
    Ordenacion = " ORDER BY codclien, codusuar "
    
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where codclien is null"

    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
        Text1(0).BackColor = vbLightBlue 'vbYellow
    End If
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
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
        
'        If EsCabecera Then 'Llama desde VerTodos del Form
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
            CadB = CadB & " and " & Aux
            
'            Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
'            cadB = cadB & " and " & Aux
'
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Text1(0).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000000")
            
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
Dim encontrado
'Form Mantenimiento de Clientes
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Clien
'    encontrado = DevuelveDesdeBD(conAri, "nomsocio", "ssocio", "codsocio", Text1(4).Text, "T")
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2)
    HaDevueltoDatos = True
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
    
End Sub

Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag) + 1
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFE_DatoSeleccionado(CadenaSeleccion As String)
'Formas de Envio
Dim indice As Byte
    indice = 29
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Envio
End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte
    indice = 14
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
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

    If frmMen.OpcionMensaje = 11 Then
        'En cadenaseleccion tenemos la WHERE que selecciona las lineas de la factura
        'que nos queremos traer para generar un albaran de rectificacion
        'Insertaremos estas lineas en la tabla slialb, y luego se podran eliminar,modificar,etc. (son de apoyo)
         InsertarLineasFactu (CadenaSeleccion)
    Else
'        If Text1(30).Text = "ART" Then
'            'Albaran de factura rectificativa
'            If Not QuitarNumSeriesAlbVenta(CadenaSeleccion) Then MsgBox "Los nº de serie a rectificar no se han actualizado correctamente.", vbExclamation
'        Else
'            If Not AsignarNumSeriesAlbVenta(CadenaSeleccion) Then
'                MsgBox "Los nº de serie del albaran no se han actualizado correctamente.", vbExclamation
'            End If
'        End If
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
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    TerminaBloquear
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            HaDevueltoDatos = False
            PonerFoco Text1(0)
            Set frmC = New frmFacClientes
            frmC.DatosADevolverBusqueda = "0|1|"
            Screen.MousePointer = vbDefault
            frmC.Show vbModal
            Set frmC = Nothing
            indice = 5
            If HaDevueltoDatos Then
                txtAnterior = ""
                Text1_LostFocus 0
                txtAnterior = Text1(0).Text
            End If
        Case 1 'NIF para cliente de Varios
            Set frmCV = New frmGesSocios
            frmCV.DatosADevolverBusqueda = "0"
            frmCV.Show vbModal
            Set frmCV = Nothing
            indice = 6
            
    End Select
    
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
    
    If Modo = 4 Then
        If Not BLOQUEADesdeFormulario(Me) Then cmdCancelar_Click
    End If
End Sub


Private Sub cmdAux_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   indice = Index + 1
   Me.cmdAux(0).Tag = Index

   PonerFormatoFecha txtAux(indice)
   If txtAux(indice).Text <> "" Then frmF.Fecha = CDate(txtAux(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtAux(indice)
End Sub


Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub


Private Sub mnBuscarTarjeta_Click()
    BotonBuscarTarjeta
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
         BotonEliminarLinea
    Else   'Eliminar Albaran
         BotonEliminar
    End If
End Sub


Private Sub mnImportarFichero_Click()
    frmImportar.Show vbModal
End Sub

Private Sub mnImprimir_Click()
'Imprimir Albaran
    BotonImprimir 45, False '45: Informe de Albaranes
End Sub


Private Sub mnLineas_Click()
    BotonMtoLineas 0, "Tarjetas"
End Sub


Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar albaran
         If BLOQUEADesdeFormulario(Me, 1) Then BotonModificar
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
    Me.Label1(51).visible = (Modo = 5) And (vEmpresa.TieneAnalitica) And SSTab1.Tab = 0
    
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
        Case 0 'Cod. cliente
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo <> 1 Then 'Modo=1 Busqueda
                    Text2(0).Text = PonerNombreDeCod(Text1(Index), conAri, "scliente", "nomclien")
                    If Text2(0).Text = "" Then
                        MsgBox "Código de cliente no existe. Reintroduzca.", vbExclamation
                    Else
                        If Modo = 3 Then
                            Text1(1).Text = SugerirCodigoSiguienteStr("scatar", "codusuar", "codclien=" & DBSet(Text1(0).Text, "N"))
                        End If
                    End If
                End If
            End If
            
        Case 1 ' usuario
            PonerFormatoEntero Text1(Index)
            
            
        Case 2 ' nombre de usuario
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    'Poner el valor del combo Tipos de Movimiento Asociado
'    If Me.cboTipomov.ListIndex <> -1 Then
'        Text1(30).Text = ObtenerCodTipom
'    End If

    CadB = ObtenerBusquedaNew(Me)  ' (Me, False)
    
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


Private Sub HacerBusquedaTarjeta()
Dim CadB As String
Dim CADENA As String
Dim i As Integer
    'Poner el valor del combo Tipos de Movimiento Asociado
'    If Me.cboTipomov.ListIndex <> -1 Then
'        Text1(30).Text = ObtenerCodTipom
'    End If

    CadB = ObtenerBusquedaNew(Me) ',  False)
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
'        If Me.EsHistorico = False Then
'            cadB = cadB & " and codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los ALV
'        End If
        CadenaConsulta = "select * from " & NombreTabla & " WHERE (codclien,codusuar) in (select codclien, codusuar from slitar where " & CadB & ")"
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
    
    Cad = Cad & ParaGrid(Text1(0), 10, "Codigo")
    Cad = Cad & "Socio|scliente|nomclien|N||40·"
    Cad = Cad & ParaGrid(Text1(1), 10, "Usuario")
    Cad = Cad & ParaGrid(Text1(2), 40, "Nombre")
    Tabla = NombreTabla & " INNER JOIN scliente ON scatar.codclien = scliente.codclien "
    Titulo = "Tarjetas"
    
    devuelve = "0|1|2|"
           
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
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 And Not BusquedaTarjetas Then
            PonerFoco Text1(kCampo)
            Text1(0).BackColor = vbLightBlue 'vbYellow
        End If
        If BusquedaTarjetas Then
            PonerFoco txtAux(0)
            txtAux(0).BackColor = vbLightBlue 'vbYellow
        End If
        Exit Sub
    Else
'            For I = 0 To 2
'                txtAux(I).visible = False
'            Next I
        
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
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo EPonerLineas

    'Limpiar campos
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i


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
    
    PonerCamposForma Me, Data1
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    Text2(0).Text = DevuelveDesdeBDNew(conAri, "scliente", "nomclien", "codclien", Text1(0).Text, "N")
    
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


    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
    
    
    
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
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
        
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
'    'Campo Nº Albaran y Tipo Movim. siempre bloqueado, excepto si estamos en modo de busqueda
'    b = (Modo <> 1)
'    BloquearTxt Text1(1), b, True
    
    b = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo <> 5) And Not BusquedaTarjetas
    Next i
    
    For i = 0 To 2
        txtAux(i).visible = (Modo = 5 And (ModificaLineas = 1 Or ModificaLineas = 2)) Or (Modo = 1 And BusquedaTarjetas) 'Not ((Modo <> 5) And Not BusquedaTarjetas)
    Next i
    
    Text1(0).Enabled = (Modo = 3) Or (Modo = 1 And Not BusquedaTarjetas)
    Text1(1).Enabled = (Modo = 3) Or (Modo = 1 And Not BusquedaTarjetas)
    Text1(2).Enabled = (Modo = 3) Or (Modo = 4) Or (Modo = 1 And Not BusquedaTarjetas)
    
    SSTab1.Enabled = (Modo = 5) Or BusquedaTarjetas
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b Or (Modo = 5 And BusquedaTarjetas)
    cmdAceptar.visible = b Or (Modo = 5 And BusquedaTarjetas)
    
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
'    Me.imgBuscar(0).visible = False
'    Me.imgBuscar(0).Enabled = (Modo = 1)
              
    Me.imgFich(0).Enabled = (Modo = 5)
    Me.imgFich(0).visible = (Modo = 5)
    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
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


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
Dim i As Byte
Dim Aux As String

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True
    
    DatosOkLinea = b
    Exit Function
    
EDatosOkLinea:
    MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: BotonVerTodos  'Todos
            
        Case 5: mnNuevo_Click 'Nuevo
        Case 6: mnModificar_Click  'Modificar
        Case 7: mnEliminar_Click  'Borrar
            
        Case 10: mnLineas_Click  'Lineas
        
        Case 11: mnBuscarTarjeta_Click ' Buscar tarjeta
        Case 12: mnImportarFichero_Click ' Importar Fichero
        
        Case 15: mnImprimir_Click ' impresion de tarjetas
            
        Case 16: mnSalir_Click   'Salir
            
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

  
'DesdeRecuperaParaRectificativa:  Para que no inserte el punto verde
Private Function InsertarLinea(numlinea As String, DesdeRecuperaParaRectificativa As Boolean) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim b As Boolean
Dim DentroTRANS As Boolean

    InsertarLinea = False
    Sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    Me.cmdAux(0).Tag = numlinea 'Aqui almaceno el Nº linea que acabo de Insertar
    
    
    If DatosOkLinea() Then 'Lineas de Albaranes
        'Inserta en tabla "slitar"
        Sql = "INSERT INTO " & NomTablaLineas
        Sql = Sql & "(codclien,codusuar,numlinea,numtarje,textoa1,textoa2,textoa3,fechaemi,fechacad,pistagr1,pistagr2,pistagr3,nomfiche) "
        Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & Val(Text1(1).Text) & ", " & numlinea & ", " & Val(txtAux(0).Text) & ","
        Sql = Sql & DBSet(txtAux(3).Text, "T") & ", " & DBSet(txtAux(4).Text, "T") & ", " & DBSet(txtAux(5).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(1).Text, "F") & ", " & DBSet(txtAux(2).Text, "F") & ", "
        Sql = Sql & DBSet(txtAux(6).Text, "T") & ", " & DBSet(txtAux(7).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(8).Text, "T") & ","
        Sql = Sql & DBSet(txtAux(9).Text, "T") & ") "
    Else
        Exit Function
    End If
    
    If Sql <> "" Then
        On Error GoTo EInsertarLinea
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute Sql
        
    
    End If
    
    conn.CommitTrans
    InsertarLinea = True
        
    Exit Function
    
EInsertarLinea:
    If Err.Number <> 0 Then
        InsertarLinea = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Albaran" & vbCrLf & Err.Description
    End If

End Function


Private Function ModificarLinea() As Boolean
'Modifica un registro en la tabla de lineas de Albaran: slialb
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
    b = True
    
    '#### LAURA 15/11/2006
    conn.BeginTrans
        
    Sql = "UPDATE " & NomTablaLineas & " Set numtarje = " & DBSet(txtAux(0).Text, "T") & ", fechaemi=" & DBSet(txtAux(1).Text, "F") & ", "
    Sql = Sql & "fechacad=" & DBSet(txtAux(2).Text, "F", "S") & ", textoa1=" & DBSet(txtAux(3).Text, "T") & ", "
    Sql = Sql & "textoa2= " & DBSet(txtAux(4).Text, "T") & ", textoa3=" & DBSet(txtAux(5).Text, "T") & ","
    Sql = Sql & "pistagr1= " & DBSet(txtAux(6).Text, "T") & ", "
    Sql = Sql & "pistagr2= " & DBSet(txtAux(7).Text, "T") & ","
    Sql = Sql & "pistagr3= " & DBSet(txtAux(8).Text, "T") & ", "
    Sql = Sql & "nomfiche=" & DBSet(txtAux(9).Text, "T", "S")
    
    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas) & " AND numlinea=" & Data2.Recordset!numlinea
    
    conn.Execute Sql
                
    
EModificarLinea:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modificar Lineas" & vbCrLf & Err.Description
        b = False
    End If
    If b Then
        conn.CommitTrans
        ModificarLinea = True
        
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
    
    CargaGridGnral DataGrid1, Me.Data2, Sql, PrimeraVez
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    b = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not b
    PrimeraVez = False
    Exit Sub
    
    
    
    
ECargaGrid:
    MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Byte
Dim tots As String

    On Error GoTo ECargaGrid

    vData.Refresh
    
    tots = "N||||0|;N||||0|;N||||0|;S|txtAux(0)|T|Tarjeta|1470|;S|txtAux(1)|T|F.Emisión|1300|;S|cmdAux(0)|B||0|;S|txtAux(2)|T|F.Caducidad|1300|;S|cmdAux(1)|B||0|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;"
    tots = tots & "N||||0|;N||||0|;N||||0|;N||||0|;"
    arregla tots, DataGrid1, Me
    
    
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
        For i = 0 To 2 'TextBox
            txtAux(i).Top = 290
            txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
        imgFich(0).visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux.Count - 1
                txtAux(i).Text = ""
                BloquearTxt txtAux(i), False
            Next i
            
        Else 'Vamos a modificar
            For i = 0 To 2
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
               
        
        'Fijamos altura(Height) y posición Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 10)
        
        For i = 0 To 2
            txtAux(i).Top = alto
            txtAux(i).Height = DataGrid1.RowHeight
        Next i
        cmdAux(0).Top = alto
        cmdAux(1).Top = alto
'        cmdAux(9).Top = alto
        cmdAux(0).Height = DataGrid1.RowHeight
        cmdAux(1).Height = DataGrid1.RowHeight
'        cmdAux(9).Height = DataGrid1.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'tarjeta
        txtAux(0).Left = DataGrid1.Left + 330
        txtAux(0).Width = DataGrid1.Columns(3).Width - 10
        'fecha de emision
        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 10
        txtAux(1).Width = DataGrid1.Columns(4).Width - 160
        cmdAux(0).Left = txtAux(1).Left + txtAux(1).Width - 50
        
        'fecha de caducidad
        txtAux(2).Left = cmdAux(0).Left + cmdAux(0).Width + 20
        txtAux(2).Width = DataGrid1.Columns(5).Width - 160
        cmdAux(1).Left = txtAux(2).Left + txtAux(2).Width - 50
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To 2
             txtAux(i).visible = visible
        Next i
        cmdAux(0).visible = visible
        cmdAux(1).visible = visible
    End If
End Sub



Private Sub txtAux_GotFocus(Index As Integer)
Dim cadkey As Integer

    cadkey = ObtenerCadKey(kCampo, Index)
    kCampo = Index
    If BusquedaTarjetas Then txtAux(Index).BackColor = vbLightBlue 'vbYellow
    ConseguirFocoLin txtAux(Index), cadkey
'    LabelAyudatxtAux Index, lblF
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

    End If
    KEYdown KeyCode
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
    
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) And Not BusquedaTarjetas Then Exit Sub
    
    If PulsadoMas2 Then
        'Para que cuando pulse el mas abra el form
        PulsadoMas2 = False
        txtAux(Index).Text = Mid(txtAux(Index).Text, 1, Len(txtAux(Index).Text) - 1)
        Exit Sub
    End If
    
    Select Case Index
        Case 0 'tarjeta
            'Comprobar que existe el almacen
'            devuelve = PonerAlmacen(txtAux(Index).Text)
'            txtAux(Index).Text = devuelve
'            If devuelve = "" Then PonerFoco txtAux(Index)

        Case 1, 2 'fecha emision y fecha de caducida
            PonerFormatoFecha txtAux(Index)
        
    End Select
    
End Sub


Private Sub BotonMtoLineas(numTab As Integer, Cad As String)

    Me.SSTab1.Tab = numTab
    TituloLinea = Cad
    ModificaLineas = 0
    
'        If vParamAplic.ArtReciclado <> "" Then
'            ClienteConTasaReciclado = Val(DevuelveDesdeBD(conAri, "tasareciclado", "sclien", "codclien", Text1(4).Text)) = 1
'        Else
'            ClienteConTasaReciclado = False
'        End If
    
    PonerModo 5
    PonerBotonCabecera True
    
    
End Sub


Private Function Eliminar(NumAlbElim As Long) As Boolean
Dim Sql As String
Dim SQL1 As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim MenError As String

    On Error GoTo FinEliminar
    conn.BeginTrans
    
    b = True
    SQL1 = ObtenerWhereCP(False)
    
    Sql = "delete from slitar where " & Replace(SQL1, "scatar", "slitar")
    conn.Execute Sql
    
    Sql = "delete from scatar where " & SQL1
    conn.Execute Sql
    
    
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
    
    Sql = " " & NombreTabla & ".codclien= " & Val(Text1(0).Text) & " and " & NombreTabla & ".codusuar= " & Val(Text1(1).Text)
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
    Sql = "SELECT slitar.codclien, slitar.codusuar, slitar.numlinea, slitar.numtarje, slitar.fechaemi, slitar.fechacad, textoa1, textoa2, textoa3, pistagr1, pistagr2, pistagr3,nomfiche "
    Sql = Sql & " FROM " & NomTablaLineas & " INNER JOIN " & NombreTabla & " ON scatar.codclien = slitar.codclien and scatar.codusuar = slitar.codusuar "
    
    If enlaza Then
        Sql = Sql & " " & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Else
        Sql = Sql & " WHERE slitar.codclien is null "
    End If
    Sql = Sql & " Order by slitar.codclien, slitar.codusuar, slitar.numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = ((Modo = 2) Or (Modo = 5 And ModificaLineas = 0))
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
        Toolbar1.Buttons(10).Enabled = b
        Me.mnLineas.Enabled = b
        
'        'Nº Series
'        Toolbar1.Buttons(11).Enabled = b And Not EsHistorico
'
'        'Generar Factura
'        'DAVID###
'        'Antes:
'        'Toolbar1.Buttons(12).Enabled = b And (CodTipoMov = "ALM" Or CodTipoMov = "ART")
'        'Ahora.  Cualquier tipo se puede generar la factura
'        Toolbar1.Buttons(12).Enabled = b
        
        'Imprimir
        Toolbar1.Buttons(15).Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
        Me.mnImprimir.Enabled = ((Modo = 2) And CodTipoMov <> "ALM")
'        Toolbar1.Buttons(14).Enabled = Toolbar1.Buttons(15).Enabled And vParamAplic.ArtPortes <> ""
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnvertodos.Enabled = Not b

        'Busqueda de tarjetas
        Toolbar1.Buttons(11).Enabled = Not b
        Me.mnBuscarTarjeta.Enabled = Not b

        'Importacion de tarjetas
        Toolbar1.Buttons(12).Enabled = Not b
        Me.mnImportarFichero.Enabled = Not b



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


Private Sub LimpiarDatosTarjeta()
Dim i As Byte

    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
End Sub
    

Private Function EliminarLinea() As Boolean
Dim vCStock As CStock
Dim Sql As String
Dim b As Boolean
Dim ImpReciclado As Single

    On Error GoTo EEliminarLinea

    EliminarLinea = False
    
    'Construir la SQL para eliminar la linea de la tabla "slialb"
    Sql = "Delete from " & NomTablaLineas & Replace(ObtenerWhereCP(True), NombreTabla, NomTablaLineas)
    Sql = Sql & " and numlinea=" & Data2.Recordset!numlinea
    
    conn.Execute Sql 'Eliminar linea
    EliminarLinea = True
    
EEliminarLinea:
     If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Tarjeta " & vbCrLf & Err.Description
    End If
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

    If Data2.Recordset.EOF Then Exit Sub

'   '[Monica]28/09/2012: guardamos en un parametro cual es la impresora por defecto
'    If vParamAplic.ImpresoraTarjetas = "" Then
'        MsgBox "No tiene indicada en parámetros cual es la impresora de Tarjetas. Revise.", vbExclamation
'        Exit Sub
'    End If
        
'    ActivaTarjeta

    frmVisReport.Informe = txtAux(9).Text
    frmVisReport.FormulaSeleccion = "{slitar.numtarje} = '" & Data2.Recordset.Fields!NUMTARJE & "'" ' & txtAux (0).Text & "'"
    frmVisReport.SelecImpresora = True
    frmVisReport.Show vbModal

'    DesactivaTarjeta

'    If Text1(0).Text = "" Then
'        MsgBox "Debe seleccionar un Albaran para Imprimir.", vbInformation
'        Exit Sub
'    End If
'
'    cadFormula = ""
'    cadParam = ""
'    cadSelect = ""
'    numParam = 0
'
'    '===================================================
'    '============ PARAMETROS ===========================
'    If (OpcionListado = 45) Then
'        If EsInformePortes Then
'            'Es el de portes
'             indRPT = 34
'        Else
'            'ALBARANES
'            If hcoCodTipoM = "ALZ" Then
'                indRPT = 29   'Albaranes B
'            ElseIf hcoCodTipoM = "ALR" Then
'                indRPT = 36
'            ElseIf hcoCodTipoM = "ALS" Then
'                indRPT = 39
'            Else
'                If EsHistorico Then
'                    indRPT = 11 'Hist. Albaranes clientes
'                Else
'                    indRPT = 10 'Albaran Clientes
'                End If
'            End If
'        End If
'    End If
'
'    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, ImpresionDirecta, pPdfRpt) Then Exit Sub
'
'    'Añadir el codigo de usuario como parametro para link con tabla Temporal (tmptiposiva) en el Report
'    'tabla temporal para el calculo del bruto total para cada tipo de IVA
'    cadParam = cadParam & "pCodUsu=" & vUsu.Codigo & "|"
'    numParam = numParam + 1
'
'    'PORTES
'    cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortes & """|"
'    numParam = numParam + 1
'
''--[Monica]arituclo de reciclado pasa a ser reutilizado como articulo de gastos de administracion
''    'PUNTO VERDE
''    cadParam = cadParam & "PuntoVerde=""" & vParamAplic.ArtReciclado & """|"
''    numParam = numParam + 1
'
'    'Si se imprimen importes y/o
''[Monica]--
''    devuelve = DevuelveDesdeBD(conAri, "albarcon", "sclien", "codclien", Text1(4).Text, "N")
'    devuelve = ""
'    If devuelve = "" Then devuelve = "0"
'    ' 0 "Todo"
'    ' 1 "Cantidad y Precio"
'    ' 2 "Cantidad"
'    cadParam = cadParam & "Albarcon=" & devuelve & "|"
'    numParam = numParam + 1
'
'
'    'Nombre fichero .rpt a Imprimir
'    If Not ImpresionDirecta Then
'        frmImprimir.NombreRPT = nomDocu
'        frmImprimir.NombrePDF = pPdfRpt
'    End If
'
'
'    '===================================================
'    '================= FORMULA =========================
'    'Cadena para seleccion Nº de Albaran
'    '---------------------------------------------------
'    If Text1(0).Text <> "" Then
'        'Cod Tipo Movimiento
'        devuelve = "{" & NombreTabla & ".codtipom}='" & CodTipoMov & "'"
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'        'Nº Albaran
'        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'        cadSelect = cadFormula
'
'        If EsHistorico Then
'            'El campo fecha tambien es clave primaria
'            devuelve = Text1(1).Text
'            devuelve = "{" & NombreTabla & ".fechaalb}=Date(" & Year(devuelve) & "," & Month(devuelve) & "," & Day(devuelve) & ")"
'            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
'
'            devuelve = "{" & NombreTabla & ".fechaalb}='" & Format(Text1(1).Text, FormatoFecha) & "'"
'            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'        End If
'
'    End If
'
'    '=========================================================================
'    'Aqui sabemos que valor tiene CodClien y añadimos a los parametros el tipo de IVA
'    'que se aplica a ese cliente
''[Monica]--
''    devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", Text1(4).Text, "N")
'    devuelve = "1"
'    If devuelve <> "" Then
'        cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
'        numParam = numParam + 1
'    End If
'
'
'    '==============================================================
'    'Comprobar si hay registros a Mostrar antes de abrir el Informe
'    devuelve = NombreTabla & " INNER JOIN " & NomTablaLineas & " ON "
'    devuelve = devuelve & NombreTabla & ".codtipom=" & NomTablaLineas & ".codtipom AND " & NombreTabla & ".numalbar= " & NomTablaLineas & ".numalbar "
'    If EsHistorico Then devuelve = devuelve & " AND " & NombreTabla & ".fechaalb= " & NomTablaLineas & ".fechaalb "
'    If Not HayRegParaInforme(devuelve, cadSelect) Then Exit Sub
'
'
'    If ImpresionDirecta Then
'        'Imrpimie directamente. Tipo 4tonda.  -----------
'        If MsgBox("¿Imprimir el albarán?", vbQuestion + vbYesNo) = vbYes Then ImprimirDirectoAlb cadSelect
'    Else
'        With frmImprimir
'            'Febrero 2010
'            If indRPT = 34 Then
'                .outTipoDocumento = 0
'            Else
'                .outTipoDocumento = 4
'                .outClaveNombreArchiv = Text1(30).Text & Text1(0).Text
'                .outCodigoCliProv = CLng(Text1(4).Text)
'            End If
'
'            .FormulaSeleccion = cadFormula
'            .OtrosParametros = cadParam
'            .NumeroParametros = numParam
'            .SoloImprimir = False
'            .EnvioEMail = False
'            .Opcion = OpcionListado
'            If indRPT = 34 Then
'                .Titulo = "Portes albaran "
'            Else
'                .Titulo = "Albaran de Socio"
'            End If
'            .ConSubinforme = True
'            .Show vbModal
'        End With
'    End If
End Sub



Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
    On Error GoTo EInsertarCab
    
'    Sql = CadenaInsertarDesdeForm(Me)
'    If Sql <> "" Then
        If InsertarDesdeForm(Me, 1) Then
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
            'Ponerse en Modo Insertar Lineas
            BotonMtoLineas 0, "Tarjetas"
            BotonAnyadirLinea
        End If
'    End If
    Text1(0).Text = Format(Text1(0).Text, "0000000")
    
'    Me.SSTab1.Tab = 0
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


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


Private Sub InsertarLineasFactu(cadWHERE)
'cadSerie = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) "
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
'        SQL = SugerirCodigoSiguienteStr("slialb", "numlinea", SQL)
'        i = Int(SQL)
    
        cadI = ""
    
        Sql = "SELECT * FROM slifac WHERE " & cadWHERE
    
        Set RS = New ADODB.Recordset
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
            txtAux(0).Text = RS!codAlmac
            txtAux(1).Text = RS!codArtic
            txtAux(2).Text = RS!NomArtic
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
            Else
                txtAux(9).Text = DBLet(RS!CodCCost, "T")
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
        
        
'        If cadI <> "" Then
'            SQL = "INSERT INTO slialb(codtipom,numalbar,numlinea,codalmac,codartic,nomartic,ampliaci,cantidad,precioar,dtoline1,dtoline2,importel,origpre) VALUES "
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




Private Sub BotonRecuperarFactura()
'Genera una factura a partir del Albaran de Mostrador
'pero sin coger contador de factura lo pide en un form

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


'Nuevo. Cuando pulse MAS (y es el primer carcater abre el prismatico asociado)
Private Sub PulsarTeclaMas(InsertandoCabecera As Boolean, Index As Integer)

    If InsertandoCabecera Then
        EsCabecera = True
        imgBuscar_Click Index
        
    Else
        'Lineas
        EsCabecera = False
        
        
    End If
        
End Sub

Private Sub CargaForaGrid()
    
    If DataGrid1.Columns.Count <= 2 Then Exit Sub
    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
    txtAux(3) = DBLet(Data2.Recordset.Fields(6).Value, "T")
    txtAux(4) = DBLet(Data2.Recordset.Fields(7).Value, "T")
    txtAux(5) = DBLet(Data2.Recordset.Fields(8).Value, "T")
    txtAux(6) = DBLet(Data2.Recordset.Fields(9).Value, "T")
    txtAux(7) = DBLet(Data2.Recordset.Fields(10).Value, "T")
    txtAux(8) = DBLet(Data2.Recordset.Fields(11).Value, "T")
    txtAux(9) = DBLet(Data2.Recordset.Fields(12).Value, "T")
'    txtAux(10) = DBLet(Data2.Recordset.Fields(13).Value, "T")

End Sub


Private Sub imgFich_Click(Index As Integer)
    CommonDialog1.InitDir = App.Path & "\Informes"
    CommonDialog1.Filter = "*.rpt|*.*"
    CommonDialog1.ShowOpen
    txtAux(9) = CommonDialog1.FileName
End Sub

''********* FUNCIONES PARA CARGAR LA IMPRESORA DE TARJETAS POR DEFECTO Y LUEGO DEJAR LA QUE INICIALMENTE TENIAN
''*************************************************************************************************************
'Private Sub ActivaTarjeta()
'    ImpresoraDefecto = Printer.DeviceName
'    XPDefaultPrinter vParamAplic.ImpresoraTarjetas
'End Sub
'
'Private Sub DesactivaTarjeta()
'    XPDefaultPrinter ImpresoraDefecto
'End Sub
'
'
''---------------- Procesos para cambio de impresora por defecto ------------------
'Private Sub XPDefaultPrinter(PrinterName As String)
'    Dim Buffer As String
'    Dim DeviceName As String
'    Dim DriverName As String
'    Dim PrinterPort As String
'    Dim R As Long
'    ' Get the printer information for the currently selected
'    ' printer in the list. The information is taken from the
'    ' WIN.INI file.
'    Buffer = Space(1024)
'    R = GetProfileString("PrinterPorts", PrinterName, "", _
'        Buffer, Len(Buffer))
'
'    ' Parse the driver name and port name out of the buffer
'    GetDriverAndPort Buffer, DriverName, PrinterPort
'
'       If DriverName <> "" And PrinterPort <> "" Then
'           SetDefaultPrinter PrinterName, DriverName, PrinterPort
'       End If
'End Sub
'
'Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
'    String, PrinterPort As String)
'
'    Dim iDriver As Integer
'    Dim iPort As Integer
'    DriverName = ""
'    PrinterPort = ""
'
'    ' The driver name is first in the string terminated by a comma
'    iDriver = InStr(Buffer, ",")
'    If iDriver > 0 Then
'
'         ' Strip out the driver name
'        DriverName = Left(Buffer, iDriver - 1)
'
'        ' The port name is the second entry after the driver name
'        ' separated by commas.
'        iPort = InStr(iDriver + 1, Buffer, ",")
'
'        If iPort > 0 Then
'            ' Strip out the port name
'            PrinterPort = Mid(Buffer, iDriver + 1, _
'            iPort - iDriver - 1)
'        End If
'    End If
'End Sub
'
'Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
'    ByVal DriverName As String, ByVal PrinterPort As String)
'    Dim DeviceLine As String
'    Dim R As Long
'    Dim L As Long
'    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
'    ' Store the new printer information in the [WINDOWS] section of
'    ' the WIN.INI file for the DEVICE= item
'    R = WriteProfileString("windows", "Device", DeviceLine)
'    ' Cause all applications to reload the INI file:
'    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
'End Sub
''------------------ Fin de los procesos relacionados con el cambio de impresora ----
