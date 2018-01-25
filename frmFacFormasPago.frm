VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFacFormasPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formas de Pago"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8430
   Icon            =   "frmFacFormasPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3390
      TabIndex        =   34
      Top             =   60
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   35
         Top             =   210
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   180
      TabIndex        =   32
      Top             =   60
      Width           =   3075
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   33
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
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
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
      Left            =   6090
      TabIndex        =   31
      Top             =   180
      Width           =   1605
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
      Left            =   5040
      MaxLength       =   5
      TabIndex        =   6
      Tag             =   "% Gastos Financieros|N|S|0|99.90|sforpa|porgasfi|#0.00|N|"
      Text            =   "Text1"
      Top             =   1860
      Width           =   765
   End
   Begin VB.Frame Frame3 
      Caption         =   "Forma de Pago por Adelantado"
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
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   8025
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
         Index           =   7
         Left            =   1965
         MaxLength       =   30
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   360
         Width           =   2685
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
         Index           =   9
         Left            =   6510
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "% Adelantado|N|S|0|99.90|sforpa|poradela|#0.00|N|"
         Text            =   "Text1"
         Top             =   360
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
         Index           =   7
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   9
         Tag             =   "Forma de Pago por adelantado|N|S|0|999|sforpa|forpapor|000|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   645
      End
      Begin VB.Image imgFPago 
         Height          =   240
         Index           =   1
         Left            =   1065
         ToolTipText     =   "Buscar forma de pago"
         Top             =   390
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "% Adelantado"
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
         Left            =   4890
         TabIndex        =   28
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "F. Pago"
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
         Left            =   270
         TabIndex        =   27
         Top             =   375
         Width           =   795
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
      Index           =   0
      Left            =   240
      MaxLength       =   15
      TabIndex        =   0
      Tag             =   "Código Forma de Pago|N|N|0|999|sforpa|codforpa|000|S|"
      Text            =   "Text1"
      Top             =   1140
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
      Index           =   4
      Left            =   3615
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "Resto Vencimientos|T|S|||sforpa|restoven||N|"
      Text            =   "Text1"
      Top             =   1860
      Width           =   1185
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
      Left            =   2175
      MaxLength       =   5
      TabIndex        =   4
      Tag             =   "Primer Vencimiento|N|N|0||sforpa|primerve|0|N|"
      Text            =   "Text1"
      Top             =   1860
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
      Index           =   2
      Left            =   240
      MaxLength       =   5
      TabIndex        =   3
      Tag             =   "Nº Vencimientos|N|N|1|99999|sforpa|numerove||N|"
      Text            =   "Text1"
      Top             =   1860
      Width           =   1005
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
      ItemData        =   "frmFacFormasPago.frx":000C
      Left            =   5040
      List            =   "frmFacFormasPago.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Tipo de Pago|N|N|||sforpa|tipforpa||N|"
      Top             =   1140
      Width           =   2700
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
      Left            =   6990
      TabIndex        =   12
      Top             =   4740
      Visible         =   0   'False
      Width           =   1135
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
      Left            =   1000
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre de la Forma de Pago|T|N|||sforpa|nomforpa|||"
      Text            =   "Text1"
      Top             =   1140
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   150
      TabIndex        =   16
      Top             =   4635
      Width           =   3135
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
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   2715
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
      Left            =   6990
      TabIndex        =   13
      Top             =   4725
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
      Left            =   5730
      TabIndex        =   11
      Top             =   4725
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   150
      Top             =   4875
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin VB.Frame Frame2 
      Caption         =   "Forma de Pago Alternativa"
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
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   8055
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
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "Forma de Pago alternativa|N|S|0|999|sforpa|forpaalt|000|N|"
         Text            =   "Text1"
         Top             =   360
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
         Left            =   6510
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Importe Mínimo|N|S|0||sforpa|impormin|#,###,###,##0.00|N|"
         Text            =   "Text1"
         Top             =   360
         Width           =   1365
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
         Index           =   6
         Left            =   1965
         MaxLength       =   30
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   360
         Width           =   2685
      End
      Begin VB.Label Label1 
         Caption         =   "F. Pago"
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
         Left            =   270
         TabIndex        =   29
         Top             =   375
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Mínimo"
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
         Left            =   4920
         TabIndex        =   25
         Top             =   375
         Width           =   1545
      End
      Begin VB.Image imgFPago 
         Height          =   240
         Index           =   0
         Left            =   1050
         Tag             =   "-1"
         ToolTipText     =   "Buscar forma de pago"
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      Caption         =   "% Gastos Financieros"
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
      Left            =   5040
      TabIndex        =   30
      Top             =   1620
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Resto Vtos."
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
      Left            =   3615
      TabIndex        =   23
      Top             =   1620
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Primer Vto."
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
      Left            =   2175
      TabIndex        =   22
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Vencimientos"
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
      TabIndex        =   21
      Top             =   1620
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Pago"
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
      Left            =   5040
      TabIndex        =   20
      Top             =   900
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "Denominación"
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
      Left            =   1005
      TabIndex        =   19
      Top             =   900
      Width           =   1695
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
      Left            =   240
      TabIndex        =   18
      Top             =   900
      Width           =   705
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
         Caption         =   "&Ver todos"
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
Attribute VB_Name = "frmFacFormasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 303

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmFPa As frmBasico2
Attribute frmFPa.VB_VarHelpID = -1

'  Variables comunes a todos los formularios
Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean


Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSCAR
            HacerBusqueda
        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    InsertarEnTesoreria
                    PonerModo 0
                End If
            End If
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario(Me, 1) Then
                    ModificarENTesoeria
                    TerminaBloquear
                    cad = "(codforpa=" & Text1(0).Text & ")"
                    If SituarData(Data1, cad, Indicador) Then
                        PonerModo 2
                        lblIndicador.Caption = Indicador
                        PonerFoco Text1(0)
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3
            LimpiarCampos
            PonerModo 0
        Case 4
            'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
    End Select
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("sforpa", "codforpa")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then 'Modo 1: Buscar
        LimpiarCampos
        PonerModo 1
        '### A mano
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
    LimpiarCampos
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index, True
    PonerCampos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    '### A mano
    'Bloquear importe minimo y %adelantado si las formas de pago estan vacias
    If Text1(6).Text = "" Then BloquearTxt Text1(8), True
    If Text1(7).Text = "" Then BloquearTxt Text1(9), True
    'Bloquear Restos Vencimientos si nº vencimientos=1
    If Val(Text1(2).Text) = 1 Then BloquearTxt Text1(4), True
    
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    If Not PuedeModificarFPenContab Then Exit Sub
    
    '### a mano
    cad = "¿Seguro que desea eliminar la Forma de Pago?" & vbCrLf
    cad = cad & vbCrLf & "Cod. Forma Pago: " & Format(Data1.Recordset.Fields(0), "000")
    cad = cad & vbCrLf & "Desc. Forma Pago: " & Data1.Recordset.Fields(1)
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        Screen.MousePointer = vbHourglass

        cad = "En aritaxi"
        Data1.Recordset.Delete
        
        
        'Para eliminar en tesoreria
        cad = "DELETE FROM sforpa WHERE codforpa = " & Text1(0).Text
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else 'Solo habia un registro
            LimpiarCampos
            PonerModo 0
        End If
        
        
        'Borro en tesoreria
        
        ConnConta.Execute cad
        
    End If
    Screen.MousePointer = vbDefault
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Forma de Pago" & vbCrLf & cad, Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(1)
End Sub

Private Sub Form_Load()
Dim Im
    'Icono del formulario
    Me.Icon = frmppal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Eliminar
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
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With

    For Each Im In Me.imgFPago
        Im.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    
    LimpiarCampos
    
    'Si hay algun combo los cargamos
    CargarComboTipoPago
    
    '## A mano
    NombreTabla = "sforpa"
    Ordenacion = " ORDER BY codforpa"
           
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '## A mano
    Data1.RecordSource = "Select * from " & NombreTabla & " where codforpa=-1"
    Data1.Refresh
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox del form
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Combo1.ListIndex = -1
End Sub


Private Sub CargarComboTipoPago()
'Cargaremos el combo, o bien desde una tabla o con valores fijos o como
'se quiera, la cuestion es cargarlo
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo
'0-Contado, 1-Cheque, 2-Pagaré, 3-Transferencia, 4-Efecto
Dim Rs As ADODB.Recordset
Dim Sql As String

    Combo1.Clear
        
    On Error GoTo ECargar

    Sql = "SELECT tipforpa, destippa from stippa"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Combo1.AddItem Rs!destippa
        Combo1.ItemData(Combo1.NewIndex) = Rs!tipforpa
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    
    
ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando combo tipos de pago.", Err.Description
    
'## ANTES
'    Combo1.AddItem "Efectivo"
'    Combo1.ItemData(Combo1.NewIndex) = 0
'
'    Combo1.AddItem "Transferencia"
'    Combo1.ItemData(Combo1.NewIndex) = 1
'
'    Combo1.AddItem "Talón"
'    Combo1.ItemData(Combo1.NewIndex) = 2
'
'    Combo1.AddItem "Pagaré"
'    Combo1.ItemData(Combo1.NewIndex) = 3
'
'    Combo1.AddItem "Recibo Bancario"
'    Combo1.ItemData(Combo1.NewIndex) = 4
'
'    Combo1.AddItem "Confirming"
'    Combo1.ItemData(Combo1.NewIndex) = 5
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
Dim indice As Integer

    If CadenaDevuelta <> "" Then
        If Val(imgFPago(0).Tag) >= 0 Then 'Llama desde Prismaticos
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgFPago(0).Tag)
            Text1(indice + 6).Text = Format(RecuperaValor(CadenaDevuelta, 1), "000")
            Text2(indice + 6).Text = RecuperaValor(CadenaDevuelta, 2)

            If Modo = 3 Then
                 Text1(indice + 8).Locked = False
                 Text1(indice + 8).BackColor = vbWhite
            End If
        Else
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
            'Sabemos que campos son los que nos devuelve
            'Creamos una cadena consulta y ponemos los datos
            CadB = ""
            Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadB = Aux
            '   Como la clave principal es unica, con poner el sql apuntando
            '   al valor devuelto sobre la clave ppal es suficiente
            'Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 2)
            'If CadB <> "" Then CadB = CadB & " AND "
            'CadB = CadB & Aux
            'Se muestran en el mismo form
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub
    
    
Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "codforpa = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgFPago_Click(Index As Integer)
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
      
    'En el Tag almacenamos el indice de la imagen de
    'Busqueda que va a llamar a frmBuscaGrid para busqueda
    imgFPago(0).Tag = Index
    MandaBusquedaPrevia ""
    imgFPago(0).Tag = -1
    PonerFoco Text1(Index + 6)
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
'    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
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
   
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod Forma de Pago
           If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
           End If
            
        Case 2 'Numero Vencimientos
            PonerFormatoEntero Text1(Index)
            If Val(Text1(Index).Text) = 1 Then
                Text1(4).Text = ""
                BloquearTxt Text1(4), True
            Else
                BloquearTxt Text1(4), False
            End If
                
        Case 3, 4 'nº vencimientos
            PonerFormatoEntero Text1(Index)
        
        Case 5, 9 '5: %Gastos Financieros, 9: %Adelantado
             'Formato tipo 4: Decimal(4,2)
             PonerFormatoDecimal Text1(Index), 4

        Case 8       '8:Importe Mínimo
            'Formato tipo 1: Decimal(12,2)
             PonerFormatoDecimal Text1(Index), 1
        
        Case 6, 7 ' 6: Forma de Pago Alternativa
                  ' 7: Forma de Pago por Adelantado
             If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "sforpa", "nomforpa", "codforpa", "N")
                If Text2(Index).Text = "" Then PonerFoco Text1(Index)
                BloquearTxt Text1(Index + 2), False
             Else
                 Text2(Index).Text = ""
                 Text1(Index + 2).Text = "" 'Importe Mínimo
                 'Modo 1: Busqueda
                 If Modo <> 1 Then BloquearTxt Text1(Index + 2), True
            End If
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then 'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
Dim cad As String

    Set frmFPa = New frmBasico2
    
    AyudaFormasPago frmFPa, , CadB
    
    Set frmFPa = Nothing

End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        If Modo = 1 Then 'Busqueda
             MsgBox "No hay ningún registro en la tabla " & NombreTabla & " para ese criterio de Búsqueda.", vbInformation
             PonerFoco Text1(0)
        Else
            MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        End If
'         MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
         Screen.MousePointer = vbDefault
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
    PonerCamposForma Me, Me.Data1
    Text2(6).Text = PonerNombreDeCod(Text1(6), 1, "sforpa", "nomforpa", "codforpa")
    Text2(7).Text = PonerNombreDeCod(Text1(7), 1, "sforpa", "nomforpa", "codforpa")
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    '--------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2) Or Kmodo = 0
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Me.Data1.Recordset.RecordCount > 1 ' Me.Toolbar1, btnPrimero, b, NumReg
    
    
    '----------------------------------------------
    'Modo insertar o modificar
    b = (Kmodo >= 3) '-->Luego not b sera kmodo<3
    cmdAceptar.visible = b Or Modo = 1
    cmdCancelar.visible = b Or Modo = 1
    If b Or Modo = 1 Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    
    BloquearText1 Me, Modo
    BloquearCmb Me.Combo1, Modo = 0 Or Modo = 2
    
    'Formas de Pago
    For I = 0 To Text2.Count - 1
        BloquearTxt Text2(I), True
    Next I
    
    Combo1.Enabled = (Modo = 3) Or (Modo = 4) Or (Modo = 1)
    
    b = (Modo = 3) 'Insertar
    'Campos Importe Mínimo y % Adelantado
    If b Then
        For I = 8 To 9
            BloquearTxt Text1(I), True
        Next I
    End If

     
    chkVistaPrevia.Enabled = (Modo <= 2)

    PonerLongCampos

    PonerModoOpcionesMenu 'Activar opciones de menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
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
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
            
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Sub PonerModoOpcionesMenu()
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    mnEliminar.Enabled = b
    
    b = (Modo >= 3)
    'Insertar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnNuevo.Enabled = Not b
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cad As String
    
    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
     
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
        
        If b Then
            If vParamAplic.ContabilidadNueva Then
                cad = DevuelveDesdeBDNew(conConta, "formapago", "codforpa", "codforpa", Text1(0), "N")
            Else
                cad = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", Text1(0), "N")
            End If
            If cad <> "" Then
                MsgBox "Esta Forma de Pago ya existe en Tesorería. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
     
    If Not b Then Exit Function
    
    'Comprobar que si nº vencimientos es 1, el campo resto vencimientos no tiene valor
    If Val(Text1(2).Text) = 1 Then
        If Not EsVacio(Text1(4)) Then
            MsgBox "El campo Resto Vencimientos no puede tener valor si NºVtos=1", vbInformation
            b = False
        End If
    End If
    If Not b Then Exit Function
     
    'Comprobar el campo resto vencimientos
    If Not EsVacio(Text1(2)) And Not EsVacio(Text1(3)) Then
        If Val(Text1(2).Text) > 1 And EsVacio(Text1(4)) Then
            MsgBox "El Campo Resto Vencimientos debe tener valor", vbInformation
            PonerFoco Text1(4)
            b = False
        End If
    End If
    If Not b Then Exit Function
    
    'Comprobar el importe Mínimo
    'Requerido si se selecciona una forma de pago alternativa
    If Not EsVacio(Text1(6)) And EsVacio(Text1(8)) Then
       MsgBox "El campo Importe Mínimo debe tener valor", vbInformation
       PonerFoco Text1(8)
       b = False
    End If
    'Verificar que el campo Importe Minimo no tiene valor si la forma de pago es vacio
    If EsVacio(Text1(6)) And Not EsVacio(Text1(8)) Then
        MsgBox "El campo Importe Mínimo no puede tener valor si no selecciona Forma de Pago.", vbInformation
        b = False
    End If
    If Not b Then Exit Function
    
    
    'Porcentaje Adelantado
    'Requerido si se selecciona una forma de pago por adelantado
    If Not EsVacio(Text1(7)) And EsVacio(Text1(9)) Then
        MsgBox "El campo % Adelantado debe tener valor", vbInformation
        PonerFoco Text1(9)
        b = False
    End If
    'Verificar que el campo %adelantado no tiene valor si la forma de pago es vacio
    If EsVacio(Text1(7)) And Not EsVacio(Text1(9)) Then
        MsgBox "El campo %Adelantado no puede tener valor si no selecciona Forma de Pago.", vbInformation
        b = False
    End If
    If Not b Then Exit Function
        
        
        
    'Comprobaciones de TESORERIA
    If Modo = 4 Then
        'Estoy modificando
        If Not PuedeModificarFPenContab Then Exit Function
    End If
    
    DatosOk = b
End Function


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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


Private Function PuedeModificarFPenContab() As Boolean
    PuedeModificarFPenContab = False
    Set miRsAux = New ADODB.Recordset
    NumRegElim = 0
    If vParamAplic.ContabilidadNueva Then
        miRsAux.Open "Select count(*) from cobros where codforpa=" & Text1(0).Text, ConnConta, adOpenForwardOnly, adLockPessimistic
    Else
        miRsAux.Open "Select count(*) from scobro where codforpa=" & Text1(0).Text, ConnConta, adOpenForwardOnly, adLockPessimistic
    End If
    If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If vParamAplic.ContabilidadNueva Then
        miRsAux.Open "Select count(*) from pagos where codforpa=" & Text1(0).Text, ConnConta, adOpenForwardOnly, adLockPessimistic
    Else
        miRsAux.Open "Select count(*) from spagop where codforpa=" & Text1(0).Text, ConnConta, adOpenForwardOnly, adLockPessimistic
    End If
    If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
    miRsAux.Close
    
    If NumRegElim > 0 Then
        If Modo = 4 Then
            If MsgBox("Existen " & NumRegElim & " vencimientos en la tesoreria con esa forma de pago. ¿Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Function
        Else
            'NO DEJO CONTINUAR
            MsgBox "Existen " & NumRegElim & " vencimientos en la tesoreria con esa forma de pago", vbExclamation
            Exit Function
        End If
    End If
    
    'añadido en al nueva contabilidad tb pueden estar en las facturas
    If vParamAplic.ContabilidadNueva Then
        miRsAux.Open "Select count(*) from factcli where codforpa=" & Text1(0).Text, ConnConta, adOpenForwardOnly, adLockPessimistic
        If Not miRsAux.EOF Then NumRegElim = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        
        miRsAux.Open "Select count(*) from factpro where codforpa=" & Text1(0).Text, ConnConta, adOpenForwardOnly, adLockPessimistic
        If Not miRsAux.EOF Then NumRegElim = NumRegElim + DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        
        If NumRegElim > 0 Then
            If Modo = 4 Then
                If MsgBox("Existen " & NumRegElim & " facturas en la contabilidad con esa forma de pago. ¿Continuar con el proceso?", vbQuestion + vbYesNo) = vbNo Then Exit Function
            Else
                'NO DEJO CONTINUAR
                MsgBox "Existen " & NumRegElim & " facturas en la contabilidad con esa forma de pago", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    'Si llega aqui puede seguir
    PuedeModificarFPenContab = True
End Function


Private Sub ModificarENTesoeria()
Dim C As String

    If vParamAplic.ContabilidadNueva Then
        C = "UPDATE formapago set nomforpa = '" & DevNombreSQL(Text1(1).Text) & "', tipforpa=" & Me.Combo1.ItemData(Combo1.ListIndex)
        C = C & ", numerove = " & DBSet(Text1(2).Text, "N") & ", primerve = " & DBSet(Text1(3).Text, "N") & ", restoven = " & DBSet(Text1(4).Text, "N")
        
    Else
        C = "UPDATE sforpa set nomforpa = '" & DevNombreSQL(Text1(1).Text) & "', tipforpa=" & Me.Combo1.ItemData(Combo1.ListIndex)
    End If
    C = C & " WHERE codforpa = " & Text1(0).Text
    ConnConta.Execute C
End Sub


Private Sub InsertarEnTesoreria()
Dim Sql As String

    On Error Resume Next
    
    If vParamAplic.ContabilidadNueva Then
        ConnConta.Execute "INSERT INTO formapago(codforpa,nomforpa, tipforpa,numerove,primerve,restoven) VALUES (" & Text1(0).Text & ",'" & DevNombreSQL(Text1(1).Text) & "'," & Me.Combo1.ItemData(Combo1.ListIndex) & "," & DBSet(Text1(2).Text, "N") & "," & DBSet(Text1(3).Text, "N") & "," & DBSet(Text1(4).Text, "N") & ")"
    Else
        ConnConta.Execute "INSERT INTO sforpa(codforpa,nomforpa, tipforpa) VALUES (" & Text1(0).Text & ",'" & DevNombreSQL(Text1(1).Text) & "'," & Me.Combo1.ItemData(Combo1.ListIndex) & ")"
    End If
    If Err.Number <> 0 Then
        MsgBox "Error insertando en tesoreria: " & vbCrLf & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub


Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub
