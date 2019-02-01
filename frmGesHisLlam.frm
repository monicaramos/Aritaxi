VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGesHisLlam 
   Caption         =   "Histórico de Llamadas."
   ClientHeight    =   11085
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   13845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "LOCALIZACION DEL SERVICIO"
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
      Height          =   4920
      Left            =   225
      TabIndex        =   86
      Top             =   855
      Width           =   13515
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
         Index           =   42
         Left            =   1410
         MaxLength       =   80
         TabIndex        =   14
         Tag             =   "Destino|T|S|||shilla|destino|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   4350
         Width           =   5055
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
         Index           =   41
         Left            =   2550
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "Puerta|T|S|||shilla|puerllama|||"
         Text            =   "ABCDEFGHIJ"
         Top             =   3045
         Width           =   1440
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
         Index           =   40
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Numero|T|S|||shilla|numllama|||"
         Text            =   "ABCDEFGHIJ"
         Top             =   3045
         Width           =   1125
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
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Nombre|T|S|||shilla|nomclien|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2175
         Width           =   5040
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
         Left            =   7920
         MaxLength       =   14
         TabIndex        =   16
         Tag             =   "Autorización|T|S|||shilla|codautor|||"
         Text            =   "Text"
         Top             =   3195
         Width           =   1905
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
         Index           =   14
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Usuario|T|S|||shilla|codusuar|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1740
         Width           =   5040
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
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Codigo cliente|N|S|||shilla|codclien|000000||"
         Text            =   "999999"
         Top             =   1350
         Width           =   825
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
         Index           =   12
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Telefono|T|S|||shilla|telefono|||"
         Text            =   "1234567890"
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Taxitronic"
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
         Height          =   645
         Left            =   6660
         TabIndex        =   101
         Top             =   4095
         Width           =   6735
         Begin VB.CheckBox Check1 
            Caption         =   "Facturado"
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
            Index           =   0
            Left            =   660
            TabIndex        =   20
            Tag             =   "Facturado|N|S|||shilla|facturad|||"
            Top             =   210
            Width           =   1395
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Abonado"
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
            Index           =   1
            Left            =   2910
            TabIndex        =   21
            Tag             =   "Abonado|N|S|||shilla|abonados|||"
            Top             =   210
            Width           =   1305
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Validado"
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
            Index           =   2
            Left            =   5025
            TabIndex        =   22
            Tag             =   "Validado|N|S|||shilla|validado|||"
            Top             =   210
            Width           =   1305
         End
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
         Left            =   11535
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "Tipo servicio|N|S|0|1|shilla|tipservi|0||"
         Top             =   3165
         Width           =   1905
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
         Index           =   1
         Left            =   2385
         TabIndex        =   100
         Text            =   "Text2"
         Top             =   945
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
         Index           =   10
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Matricula|T|S|||shilla|matricul|||"
         Text            =   "Text"
         Top             =   3645
         Width           =   1905
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
         Left            =   11535
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "Identificacion|T|S|||shilla|idservic|||"
         Top             =   2730
         Width           =   1905
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
         Left            =   2700
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Num vehiculo|N|N|||shilla|numeruve|#####0|S|"
         Text            =   "Text"
         Top             =   510
         Width           =   1185
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
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   1
         Tag             =   "Hora|H|N|||shilla|hora|hh:mm:ss|S|"
         Text            =   "99:99:99"
         Top             =   510
         Width           =   1005
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
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   15
         Tag             =   "Licencia|T|S|||shilla|licencia|||"
         Text            =   "Text"
         Top             =   2730
         Width           =   1905
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
         Left            =   1410
         MaxLength       =   35
         TabIndex        =   13
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   3915
         Width           =   5040
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
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   12
         Tag             =   "Población|T|S|||shilla|ciudadre|||"
         Text            =   "ABCDEFGHIJKLMNÑOPQRSTUVWXYZABC"
         Top             =   3480
         Width           =   5040
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
         Left            =   5535
         MaxLength       =   6
         TabIndex        =   11
         Text            =   "HHHHHH"
         Top             =   3045
         Width           =   915
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
         Left            =   1410
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "Domicilio|T|S|||shilla|dirllama|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2610
         Width           =   5040
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
         Left            =   210
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Fecha|F|N|||shilla|fecha|dd/mm/yyyy|S|"
         Text            =   "99/99/9999"
         Top             =   510
         Width           =   1275
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
         Left            =   1395
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Codigo socio|N|N|||shilla|codsocio|000000||"
         Text            =   "Text"
         Top             =   945
         Width           =   915
      End
      Begin VB.Frame FrameAux2 
         Height          =   2415
         Left            =   6570
         TabIndex        =   87
         Top             =   180
         Width           =   6885
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   7
            Left            =   5550
            TabIndex        =   97
            Tag             =   "Documento|T|N|||shilla_fichdocs|docum|||"
            Text            =   "docum"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   6
            Left            =   4800
            TabIndex        =   96
            Tag             =   "Campo|T|N|||shilla_fichdocs|campo||S|"
            Text            =   "campo"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   5
            Left            =   4080
            TabIndex        =   95
            Tag             =   "Descripcion|T|N|||shilla_fichdocs|descripfich||N|"
            Text            =   "descripcion"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   4
            Left            =   3330
            TabIndex        =   94
            Tag             =   "Orden|N|N|0||shilla_fichdocs|orden||S|"
            Text            =   "Orden"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   3
            Left            =   2580
            TabIndex        =   93
            Tag             =   "vehiculo|N|N|0||shilla_fichdocs|numeruve||S|"
            Text            =   "Codigo"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   2
            Left            =   1800
            TabIndex        =   92
            Tag             =   "Vehiculo|N|N|0||shilla_fichdocs|numeruve||S|"
            Text            =   "Vehiculo"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   1
            Left            =   1110
            TabIndex        =   91
            Tag             =   "Hora|H|N|||shilla_fichdocs|hora|hh:mm|S|"
            Text            =   "Hora"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox txtaux3 
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
            Height          =   350
            Index           =   0
            Left            =   330
            TabIndex        =   90
            Tag             =   "Fecha|F|S|||shilla_fichdocs|fecha|dd/mm/yyyy|S|"
            Text            =   "Fecha"
            Top             =   810
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Height          =   345
            Left            =   90
            TabIndex        =   88
            Top             =   120
            Width           =   2625
            Begin MSComctlLib.Toolbar ToolbarAux0 
               Height          =   330
               Left            =   165
               TabIndex        =   89
               Top             =   45
               Width           =   2235
               _ExtentX        =   3942
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
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.ToolTipText     =   "Modificar"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Eliminar"
                  EndProperty
               EndProperty
            End
         End
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   375
            Left            =   2910
            Top             =   630
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
         Begin MSComctlLib.ListView Lw1 
            Height          =   1635
            Left            =   90
            TabIndex        =   98
            Top             =   600
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   2884
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
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
         Begin MSComDlg.CommonDialog cd1 
            Left            =   2115
            Top             =   630
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Documentos Asociados"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   99
            Top             =   210
            Width           =   3480
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Destino"
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
         Left            =   210
         TabIndex        =   120
         Top             =   4350
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Nº/Puerta"
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
         Left            =   210
         TabIndex        =   119
         Top             =   3045
         Width           =   1545
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
         Index           =   17
         Left            =   210
         TabIndex        =   118
         Top             =   2175
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Autorización"
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
         Left            =   6660
         TabIndex        =   117
         Top             =   3195
         Width           =   1485
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
         Index           =   15
         Left            =   210
         TabIndex        =   116
         Top             =   1740
         Width           =   1305
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1170
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1110
         Tag             =   "-1"
         ToolTipText     =   "Buscar Socio"
         Top             =   990
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
         Index           =   14
         Left            =   4635
         TabIndex        =   115
         Top             =   1380
         Width           =   705
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
         Index           =   13
         Left            =   210
         TabIndex        =   114
         Top             =   1380
         Width           =   1425
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   1650
         TabIndex        =   113
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   240
         TabIndex        =   112
         Top             =   285
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   5385
         Tag             =   "-1"
         ToolTipText     =   "Buscar cliente"
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Matrícula"
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
         Left            =   6660
         TabIndex        =   111
         Top             =   3690
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación"
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
         Left            =   9945
         TabIndex        =   110
         Top             =   2730
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Vehículo"
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
         Left            =   2700
         TabIndex        =   109
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de servicio"
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
         Left            =   9900
         TabIndex        =   108
         Top             =   3165
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Licencia"
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
         Left            =   6660
         TabIndex        =   107
         Top             =   2730
         Width           =   1245
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
         Index           =   7
         Left            =   210
         TabIndex        =   106
         Top             =   3915
         Width           =   1425
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
         Index           =   6
         Left            =   210
         TabIndex        =   105
         Top             =   3480
         Width           =   1545
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   5265
         Tag             =   "-1"
         ToolTipText     =   "Buscar población"
         Top             =   3075
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "CP"
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
         Left            =   4905
         TabIndex        =   104
         Top             =   3045
         Width           =   255
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
         Index           =   0
         Left            =   210
         TabIndex        =   103
         Top             =   2610
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
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
         Left            =   195
         TabIndex        =   102
         Top             =   1005
         Width           =   645
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1515
      Left            =   225
      TabIndex        =   82
      Top             =   8790
      Width           =   13515
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
         Index           =   39
         Left            =   2670
         MaxLength       =   200
         TabIndex        =   49
         Tag             =   "Observaciones 2|T|S|||shilla|observa2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   1020
         Width           =   10665
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
         Index           =   38
         Left            =   2670
         MaxLength       =   200
         TabIndex        =   48
         Tag             =   "Observaciones Cliente|T|S|||shilla|observa1|||"
         Text            =   $"frmGesHisLlam.frx":0000
         Top             =   630
         Width           =   10665
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
         Index           =   37
         Left            =   2670
         MaxLength       =   60
         TabIndex        =   47
         Tag             =   "Observaciones|T|S|||shilla|observac2|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   240
         Width           =   10665
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   2400
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   1020
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   2400
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   630
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   2400
         Tag             =   "-1"
         ToolTipText     =   "Ver Observaciones"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones II"
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
         Left            =   120
         TabIndex        =   85
         Top             =   1020
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones Cliente"
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
         Left            =   120
         TabIndex        =   84
         Top             =   630
         Width           =   2325
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
         Index           =   33
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   1785
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "IMPORTES FACTURADOS"
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
      Height          =   2940
      Left            =   6780
      TabIndex        =   71
      Top             =   5760
      Width           =   6960
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
         Index           =   43
         Left            =   2025
         TabIndex        =   39
         Tag             =   "Empresa alfa|N|S|||shilla|empresa|####0||"
         Text            =   "Text"
         Top             =   1770
         Width           =   1245
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
         Left            =   5400
         TabIndex        =   44
         Tag             =   "Imp.Propina|N|S|||shilla|imppropi|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1770
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Facturado Cliente"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   46
         Tag             =   "Facturado Cliente|N|N|0|1|shilla|facturadocliente|||"
         Top             =   2310
         Width           =   2145
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Liquidado Socio"
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
         Index           =   3
         Left            =   1080
         TabIndex        =   45
         Tag             =   "Liquidado Socio|N|N|0|1|shilla|liquidadosocio|||"
         Top             =   2310
         Width           =   1905
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
         Left            =   5400
         TabIndex        =   43
         Tag             =   "Imp.Peaje|N|S|||shilla|imppeaje|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1410
         Width           =   1245
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
         Index           =   34
         Left            =   5400
         TabIndex        =   42
         Tag             =   "Suplemento|N|S|||shilla|suplemen|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1050
         Width           =   1245
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
         Index           =   33
         Left            =   5400
         TabIndex        =   41
         Tag             =   "Distancia|N|S|||shilla|distanci|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   690
         Width           =   1245
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
         Index           =   32
         Left            =   5400
         TabIndex        =   40
         Tag             =   "Ext.Venta|N|S|||shilla|extventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   330
         Width           =   1245
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
         Index           =   31
         Left            =   2040
         TabIndex        =   38
         Tag             =   "Ext.Compra|N|S|||shilla|extcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1410
         Width           =   1245
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
         Left            =   2040
         TabIndex        =   37
         Tag             =   "Imp.Venta|N|S|||shilla|impventa|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   1050
         Width           =   1245
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
         Left            =   2040
         TabIndex        =   36
         Tag             =   "Imp.Compra|N|S|||shilla|impcompr|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   690
         Width           =   1245
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
         Index           =   28
         Left            =   2040
         TabIndex        =   35
         Tag             =   "Imp.TX|N|S|||shilla|importtx|#,###,###,##0.00||"
         Text            =   "Text"
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa Alfa"
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
         Left            =   225
         TabIndex        =   81
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Propina"
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
         Index           =   32
         Left            =   3840
         TabIndex        =   80
         Top             =   1770
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Peaje"
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
         Index           =   31
         Left            =   3840
         TabIndex        =   79
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Suplemento"
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
         Left            =   3840
         TabIndex        =   78
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Distancia"
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
         Left            =   3840
         TabIndex        =   77
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Extra Venta"
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
         Left            =   3840
         TabIndex        =   76
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Extra Compra"
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
         TabIndex        =   75
         Top             =   1410
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Venta"
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
         Left            =   240
         TabIndex        =   74
         Top             =   1050
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Compra"
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
         Index           =   23
         Left            =   240
         TabIndex        =   73
         Top             =   690
         Width           =   1845
      End
      Begin VB.Label Label1 
         Caption         =   "Importe TX"
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
         Left            =   240
         TabIndex        =   72
         Top             =   330
         Width           =   1605
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2910
      Left            =   225
      TabIndex        =   63
      Top             =   5790
      Width           =   6495
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
         Index           =   27
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   34
         Tag             =   "Operador Despa.|T|S|||shilla|opedespa|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2460
         Width           =   5205
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
         Index           =   26
         Left            =   1200
         MaxLength       =   30
         TabIndex        =   33
         Tag             =   "Operador Reserva|T|S|||shilla|opereser|||"
         Text            =   "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
         Top             =   2055
         Width           =   5205
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
         Left            =   3060
         MaxLength       =   8
         TabIndex        =   32
         Tag             =   "Hora Final|H|S|||shilla|horfinal|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1635
         Width           =   1005
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
         Left            =   3060
         MaxLength       =   8
         TabIndex        =   30
         Tag             =   "Hora Ocupado|H|S|||shilla|horocupa|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   1275
         Width           =   1005
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
         Left            =   3060
         MaxLength       =   8
         TabIndex        =   28
         Tag             =   "Hora Llegada|H|S|||shilla|horllega|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   915
         Width           =   1005
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
         Left            =   3060
         MaxLength       =   8
         TabIndex        =   26
         Tag             =   "Hora Aviso|H|S|||shilla|horaviso|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   555
         Width           =   1005
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   31
         Tag             =   "Fecha Final|F|S|||shilla|fecfinal|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1635
         Width           =   1305
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   29
         Tag             =   "Fecha ocupado|F|S|||shilla|fecocupa|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   1275
         Width           =   1305
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   27
         Tag             =   "Fecha Llegada|F|S|||shilla|fecllega|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   915
         Width           =   1305
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
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   25
         Tag             =   "Fecha Aviso|F|S|||shilla|fecaviso|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   555
         Width           =   1305
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
         Left            =   3060
         MaxLength       =   8
         TabIndex        =   24
         Tag             =   "Hora Reserva|H|S|||shilla|horreser|hh:mm:ss||"
         Text            =   "99:99:99"
         Top             =   195
         Width           =   1005
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
         Index           =   11
         Left            =   1650
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Fecha Reserva|F|S|||shilla|fecreser|dd/mm/yyyy||"
         Text            =   "99/99/9999"
         Top             =   195
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Ope.Des"
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
         Left            =   240
         TabIndex        =   70
         Top             =   2460
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Ope.Res"
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
         TabIndex        =   69
         Top             =   2055
         Width           =   1095
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1320
         Picture         =   "frmGesHisLlam.frx":00CC
         ToolTipText     =   "Buscar fecha"
         Top             =   1635
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1320
         Picture         =   "frmGesHisLlam.frx":0157
         ToolTipText     =   "Buscar fecha"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1320
         Picture         =   "frmGesHisLlam.frx":01E2
         ToolTipText     =   "Buscar fecha"
         Top             =   915
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1320
         Picture         =   "frmGesHisLlam.frx":026D
         ToolTipText     =   "Buscar fecha"
         Top             =   555
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Finalizado"
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
         Left            =   240
         TabIndex        =   68
         Top             =   1635
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Ocupado"
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
         Left            =   240
         TabIndex        =   67
         Top             =   1275
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Llegada"
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
         Left            =   240
         TabIndex        =   66
         Top             =   915
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Aviso"
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
         Left            =   240
         TabIndex        =   65
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Reserva"
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
         TabIndex        =   64
         Top             =   195
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmGesHisLlam.frx":02F8
         ToolTipText     =   "Buscar fecha"
         Top             =   195
         Width           =   240
      End
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3915
      TabIndex        =   60
      Top             =   60
      Width           =   750
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   61
         Top             =   180
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Marcar Validados"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4725
      TabIndex        =   58
      Top             =   60
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   59
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
      Left            =   240
      TabIndex        =   56
      Top             =   60
      Width           =   3615
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   57
         Top             =   180
         Width           =   3195
         _ExtentX        =   5636
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
      Left            =   8460
      TabIndex        =   55
      Top             =   210
      Width           =   1605
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1500
      Left            =   5400
      Top             =   9135
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2646
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
      Left            =   11400
      TabIndex        =   50
      Top             =   10530
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
      Left            =   12630
      TabIndex        =   51
      Top             =   10530
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
      Left            =   12630
      TabIndex        =   52
      Top             =   10530
      Visible         =   0   'False
      Width           =   1135
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   240
      TabIndex        =   53
      Top             =   10320
      Width           =   3975
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
         Left            =   120
         TabIndex        =   54
         Top             =   210
         Width           =   3615
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   1500
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2646
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
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   38
      Left            =   4320
      TabIndex        =   62
      Top             =   10530
      Visible         =   0   'False
      Width           =   2955
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
Attribute VB_Name = "frmGesHisLlam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 317

Public FechaServ As String
Public HoraServ As String
Public NumerUve As String

Public WithEvents frmLlamPre As frmGesHisLlamPrev
Attribute frmLlamPre.VB_VarHelpID = -1
Public WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1
Public WithEvents frmV As frmGesVehic
Attribute frmV.VB_VarHelpID = -1
Public WithEvents frmCal As frmCal
Attribute frmCal.VB_VarHelpID = -1
Public WithEvents frmS As frmGesSocios
Attribute frmS.VB_VarHelpID = -1
Public WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1

Private frmDocs   As frmFichaTecIMG  'frmDocImgs  mto de imagenes

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

'variables control del log
Dim Socio As Currency
Dim Tfno As String
Dim clien As String
Dim Ident As String
Dim Autor As String
Dim TipoSer As Byte
Dim Licencia As String
Dim Matricula As String
Dim TxFactu As Byte
Dim TxAbo As Byte
Dim TxVali As Byte
Dim ImpCompra As Currency
Dim ImpVta As Currency
Dim LiqSoc As Byte
Dim FacCli As Byte
Dim Usuario As String
Dim Nombre As String
Dim Domicilio As String
Dim Destino As String
Dim numero As String
Dim Puerta As String
Dim Poblacion As String
Dim Obser As String
Dim Obs1 As String
Dim Obs2 As String


Dim cadB1 As String
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
Private BuscaChekc As String


Dim CadB As String

' VARIABLES DOCUMENTOS ASOCIADOS
Dim It As ListItem
Dim Contador As Integer
Dim Fichero As String
Dim TipoDocu As Byte
Dim FicheroAEliminar As String

Private Const CarpetaIMG = "temp" 'ImgFicFT2



Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then CheckCadenaBusqueda Check1(Index), BuscaChekc
End Sub


Private Sub Check1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 0 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub


Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'Private Sub cmdAccCRM_Click(Index As Integer)
'    Select Case Index
'        Case 0
'            Set frmDocs = New frmFichaTecIMG
'
'            frmDocs.vDatos = Text1(1).Text & "|" & Text1(7).Text & "|" & Text1(8).Text & "|"
'            frmDocs.Opcion = Index
'            frmDocs.Show vbModal
'
'            Set frmDocs = Nothing
'
'            CargaDatosLW
'
'        Case 1 'Impresión del documento
'            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
'
'            ImprimirImagen
'
'        Case 2 'Eliminar
'            If Me.lw1.SelectedItem Is Nothing Then Exit Sub
'
'            EliminarImagen
'    End Select
'End Sub


Private Sub ImprimirImagen()
Dim NFic As Long
Dim vAdobe As String

   CargarArchivos
                
   Call ShellExecute(Me.hwnd, "Open", Me.lw1.SelectedItem.SubItems(2), "", "", 1)
   
End Sub


'Private Sub EliminarImagen()
'Dim SQL As String
'Dim Mens As String
'
'    On Error GoTo eEliminarImagen
'
'    Mens = "Va a proceder a eliminar la imágen de la lista correspondiente al servicio de fecha " & Text1(1).Text & " " & Text1(7).Text & " licencia " & Text1(8).Text & ". " & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf & vbCrLf
'
'    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'        SQL = "delete from rfichdocs where fecha = " & DBSet(Text1(1).Text, "F") & " and hora = " & DBSet(Text1(7).Text, "H") & " and numeruve = " & DBSet(Text1(8).Text, "N") & " and codigo = " & Me.lw1.SelectedItem.SubItems(3)
'        conn.Execute SQL
'
'        CargaDatosLW
'
'    End If
'    Exit Sub
'
'eEliminarImagen:
'    MuestraError Err.Number, "Eliminar imágen", Err.Description
'End Sub
'

Private Sub cmdAceptar_Click()
Dim I As Long
Dim CadB As String
Dim Cad As String
Dim Indicador As String
Dim cad1 As String

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
                     '[Monica]04/02/2015: guardo en el slog los campos que me han cambiado
                    cad1 = ""
                    'Socio
                    If Text1(0).Text <> Socio Then cad1 = cad1 & "Socio: " & Format(Socio, "000000") & " a " & Text1(0).Text & ";"
                    'Liquidado socio
                    If Check1(3).Value <> LiqSoc Then cad1 = cad1 & "LiqSoc: " & LiqSoc & " a " & Check1(3).Value & ";"
                    'Facturado Cliente
                    If Check1(4).Value <> FacCli Then cad1 = cad1 & "FacCli: " & FacCli & " a " & Check1(4).Value & ";"
                    'Taxitronic facturado
                    If Check1(0).Value <> TxFactu Then cad1 = cad1 & "TxFra: " & TxFactu & " a " & Check1(0).Value & ";"
                    'Taxitronic abonado
                    If Check1(1).Value <> TxAbo Then cad1 = cad1 & "TxAbo: " & TxAbo & " a " & Check1(1).Value & ";"
                    'Taxitronic validado
                    If Check1(2).Value <> TxVali Then cad1 = cad1 & "TxVal: " & TxVali & " a " & Check1(2).Value & ";"
                    'Tfno
                    If Trim(Text1(12).Text) <> Trim(Tfno) Then cad1 = cad1 & "Tfno: " & Trim(Tfno) & " a " & Trim(Text1(12).Text) & ";"
                    'Codclien
                    If Text1(13).Text <> clien Then cad1 = cad1 & "Cli: " & clien & " a " & Text1(13).Text & ";"
                    'Usuario
                    If Trim(Text1(14).Text) <> Trim(Usuario) Then cad1 = cad1 & "Usu: " & Trim(Usuario) & " a " & Trim(Text1(14).Text) & ";"
                    'Identificacion
                    If Trim(Text1(9).Text) <> Trim(Ident) Then cad1 = cad1 & "Iden: " & Trim(Ident) & " a " & Trim(Text1(9).Text) & ";"
                    'Autorizacion
                    If Trim(Text1(15).Text) <> Trim(Autor) Then cad1 = cad1 & "Autor: " & Trim(Autor) & " a " & Trim(Text1(15).Text) & ";"
                    'Tipo de Servicio
                    If Combo1.ListIndex <> TipoSer Then cad1 = cad1 & "TSer: " & TipoSer & " a " & Combo1.ListIndex & ";"
                    'Licencia
                    If Trim(Text1(6).Text) <> Trim(Licencia) Then cad1 = cad1 & "Lic: " & Trim(Licencia) & " a " & Trim(Text1(6).Text) & ";"
                    'Matricula
                    If Trim(Text1(10).Text) <> Trim(Matricula) Then cad1 = cad1 & "Mat: " & Trim(Matricula) & " a " & Trim(Text1(10).Text) & ";"
                    'Importe de Compra
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(29).Text))) <> ImpCompra Then cad1 = cad1 & "ImpCom: " & ImpCompra & " a " & Text1(29).Text & ";"
                    'Importe de Venta
                    If CCur(ImporteSinFormato(ComprobarCero(Text1(30).Text))) <> ImpVta Then cad1 = cad1 & "ImpVta: " & ImpVta & " a " & Text1(30).Text & ";"
                    'Nombre
                    If Trim(Text1(16).Text) <> Trim(Nombre) Then cad1 = cad1 & "Nombre: " & Trim(Nombre) & " a " & Trim(Text1(16).Text) & ";"
                    'Domicilio
                    If Trim(Text1(2).Text) <> Trim(Domicilio) Then cad1 = cad1 & "Dom: " & Trim(Domicilio) & " a " & Trim(Text1(2).Text) & ";"
                    'Numero
                    If Trim(Text1(40).Text) <> Trim(numero) Then cad1 = cad1 & "Nro: " & Trim(numero) & " a " & Trim(Text1(40).Text) & ";"
                    'Puerta
                    If Trim(Text1(41).Text) <> Trim(Puerta) Then cad1 = cad1 & "Pta: " & Trim(Puerta) & " a " & Trim(Text1(41).Text) & ";"
                    'Poblacion
                    If Trim(Text1(4).Text) <> Trim(Poblacion) Then cad1 = cad1 & "Pob: " & Trim(Poblacion) & " a " & Trim(Text1(4).Text) & ";"
                    'Destino
                    If Trim(Text1(42).Text) <> Trim(Destino) Then cad1 = cad1 & "Des: " & Trim(Destino) & " a " & Trim(Text1(42).Text) & ";"
                    
                    'Observaciones
                    If Trim(Text1(37).Text) <> Trim(Obser) Then cad1 = cad1 & "Obs: " & Trim(Obser) & " a " & Trim(Text1(37).Text) & ";"
                    'Observaciones 1
                    If Trim(Text1(38).Text) <> Trim(Obs1) Then cad1 = cad1 & "Obs1: " & Trim(Obs1) & " a " & Trim(Text1(38).Text) & ";"
                    'Observaciones 2
                    If Trim(Text1(39).Text) <> Trim(Obs2) Then cad1 = cad1 & "Obs2: " & Trim(Obs2) & " a " & Trim(Text1(39).Text) & ";"
                    
                    
                    
                                       
                    Set LOG = New cLOG
                    LOG.Insertar 9, vUsu, "Llamada modificada: " & Text1(1).Text & " " & Text1(7).Text & " " & Text1(8).Text & " " & cad1 & vbCrLf
                    Set LOG = Nothing
                
                    TerminaBloquear
                    
                    '[Monica] 02/06/2011: tras modificar volvemos al registro correspondiente
                    NumRegElim = Adodc1.Recordset.AbsolutePosition
                    Me.Adodc1.Refresh
                    If SituarDataPosicion(Adodc1, NumRegElim, "") Then
                        PonerCampos
                    End If
                    PonerModo 2

                    'fin
'[Monica] 02/06/2011: comentado
'                    PosicionarData
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
Error1:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
'    cadB1 = ObtenerBusqueda(Me, True, BuscaChekc)
    cadB1 = ObtenerBusqueda(Me, True, BuscaChekc)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(1)
    End If
End Sub
Private Function DatosOk() As Boolean

DatosOk = False
If Modo = 4 Then
    DatosOk = True
    Exit Function
End If
'CODIGO DE socio
If Text1(0).Text = "" Then
    MsgBox "Debe introducir el código de socio.", vbExclamation
    PonerFoco Text1(0)
    Exit Function
ElseIf Not IsNumeric(Text1(0).Text) Then
        MsgBox "El código de socio debe ser numérico.", vbExclamation
        PonerFoco Text1(0)
        Exit Function
End If

'Fecha
If Text1(1).Text = "" Then
    MsgBox "Debe introducir la fecha de la llamada.", vbExclamation
    PonerFoco Text1(1)
    Exit Function
ElseIf Not IsDate(Text1(1).Text) Then
    MsgBox "La fecha no tiene formato correcto.", vbExclamation
    PonerFoco Text1(1)
    Exit Function
End If

'Hora
If Text1(7).Text = "" Then
    MsgBox "Debe introducir la hora de la llamada.", vbExclamation
    PonerFoco Text1(7)
    Exit Function
ElseIf Not IsDate(Text1(7).Text) Then
    MsgBox "La hora no tiene formato correcto.", vbExclamation
    PonerFoco Text1(7)
    Exit Function
End If

'numero de vehiculo
If Text1(8).Text = "" Then
    MsgBox "Debe introducir el número de vehiculo.", vbExclamation
    PonerFoco Text1(8)
    Exit Function
ElseIf Not IsNumeric(Text1(8).Text) Then
        MsgBox "El número de vehiculo debe ser numérico.", vbExclamation
        PonerFoco Text1(8)
        Exit Function
End If
DatosOk = True
    
End Function

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim Cad As String

    'Quitar lineas y volver a la cabecera
        If Adodc1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        Cad = Adodc1.Recordset.Fields(0) & "|"
        Cad = Cad & Adodc1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(Cad)
        Unload Me
    
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()

If Combo1.Text = "0" Then
    Text1(13).Text = ""
    Text1(13).BackColor = &H80000018
    Text1(13).Enabled = False
    imgBuscar(1).Enabled = False
Else
    Text1(13).Enabled = True
    Text1(13).BackColor = &H80000005
    imgBuscar(1).Enabled = True
End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If PrimeraVez Then
        PrimeraVez = False
        If FechaServ <> "" Then
            If Me.Adodc1.Recordset.EOF Then
                PonerCadenaBusqueda
            Else
                PonerCampos
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del form
    Me.Icon = frmppal.Icon
    

'    'ICONITOS DE LA BARRA
'    btnAnyadir = 5
'    btnPrimero = 12
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(5).Image = 3   'Insertar Nuevo
'        .Buttons(6).Image = 4   'Modificar
'        .Buttons(7).Image = 5   'Borrar
'        '.Buttons(9).Image = 10 'Lineas
'        .Buttons(9).Image = 16 'Imprmir
'        .Buttons(10).Image = 15 'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
    With Me.Toolbar1
        btnPrimero = 11
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Eliminar
        .Buttons(8).Image = 16 'Imprmir
    End With
    
    With Me.Toolbar5
        .ImageList = frmppal.imgListComun1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .Buttons(1).Image = 39  'marcar como validados
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
    
    With Me.ToolbarAux0
        .HotImageList = frmppal.imgListComun_OM16
        .DisabledImageList = frmppal.imgListComun_BN16
        .ImageList = frmppal.imgListComun16
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
    End With
    
    
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 12
'    End With
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        imgBuscar(I).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    For I = 0 To Me.imgFecha.Count - 1
        imgFecha(I).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next


    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
    CargarCombo
    
    NombreTabla = "shilla"
    CadenaConsulta = "Select * from " & NombreTabla
    
    If FechaServ <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafaccli1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura
        If cadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & cadB1
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        'If Not publicidad Then
        CadenaConsulta = CadenaConsulta & " WHERE false " 'fecha is null  "
    End If
    
    '## A mano
    Ordenacion = " ORDER BY fecha,hora"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = CadenaConsulta ' "Select * from " & NombreTabla & " where numeruve=-1"
    Adodc1.Refresh
    
    
    
    CargarColumnas
    CargarDocumentos
    
    If FechaServ = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            PonerModo 1
        End If
    Else
        If Adodc1.Recordset.EOF Then
            PonerModo 0
        Else
            PonerModo 2
        End If
    End If
End Sub

Private Sub CargarCombo()
    Combo1.AddItem "Normal"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Crédito"
    Combo1.ItemData(Combo1.NewIndex) = 1
    
End Sub

Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I

    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    BuscaChekc = ""
    
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
    If Not Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible b And Me.Adodc1.Recordset.RecordCount > 1 ' Me.Toolbar1, btnPrimero, b, NumReg
    
    
    
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
    BloquearCmb Combo1, (Modo = 0 Or Modo = 2)
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    Combo1.Enabled = b
    For I = 0 To 2
        Check1(I).Enabled = b
    Next I
    
'[Monica]04/02/2015: dejamos modificar liquidado y facturado, antes solo podian consultar
    For I = 3 To 4
        Check1(I).Enabled = b '(Modo = 1)
    Next I
    
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
    
    ' No hay icono para las observaciones de 60 de longitud maxima
    Me.imgBuscar(2).Enabled = False
    Me.imgBuscar(2).visible = False
    
    Me.imgBuscar(4).Enabled = (Modo > 0)
    Me.imgBuscar(4).visible = (Modo > 0)
    Me.imgBuscar(5).Enabled = (Modo > 0)
    Me.imgBuscar(5).visible = (Modo > 0)
    
    'lineas de documentos
    b = (Modo = 5) 'And (NumTabMto = 0) And (ModoLineas <> 3)
    For I = 0 To txtAux3.Count - 1
        If (I >= 0 And I <= 3) Or (I >= 6 And I <= 7) Then
            txtAux3(I).Enabled = False
            txtAux3(I).visible = False
        Else
            txtAux3(I).Enabled = b
            txtAux3(I).visible = b
        End If
    Next I
    
    
    
    For I = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(I).Enabled = b
    Next I
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    '-----------------------------
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu según modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
    PonerModoUsuarioGnral Modo, "aritaxi"
    CargarDocumentos
    
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerModoUsuarioGnral(Modo As Byte, Aplicacion As String)
Dim Rs As ADODB.Recordset
Dim Cad As String
    
    On Error Resume Next

    Cad = "select ver, creareliminar, modificar, imprimir, especial from menus_usuarios where aplicacion = " & DBSet(Aplicacion, "T")
    Cad = Cad & " and codigo = " & DBSet(IdPrograma, "N") & " and codusu = " & DBSet(vUsu.Id, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Toolbar1.Buttons(1).Enabled = Toolbar1.Buttons(1).Enabled And DBLet(Rs!creareliminar, "N")
        Toolbar1.Buttons(2).Enabled = Toolbar1.Buttons(2).Enabled And DBLet(Rs!Modificar, "N")
        Toolbar1.Buttons(3).Enabled = Toolbar1.Buttons(3).Enabled And DBLet(Rs!creareliminar, "N")
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
    End If
    
    
    ToolbarAux0.Buttons(1).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.Adodc1.Recordset.RecordCount > 0)
    ToolbarAux0.Buttons(3).Enabled = DBLet(Rs!creareliminar, "N") And (Modo = 2 And Me.Adodc1.Recordset.RecordCount > 0)
    
    
    
    Rs.Close
    Set Rs = Nothing
    
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



Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub
Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean
Dim I As Byte

    b = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(1).Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    'imprimir
    Toolbar1.Buttons(8).Enabled = b
    '------------------------------------------
    
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = Not b
End Sub

Private Sub LimpiarCampos()
Dim I As Integer

On Error Resume Next
    
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    Me.Combo1.ListIndex = -1
    
    For I = 0 To Check1.Count - 1
        Check1(I).Value = 0
    Next I
    
    '### a mano
    If Err.Number <> 0 Then Err.Clear

End Sub


Private Sub frmCal_Selec(vFecha As Date)
    Fecha = vFecha
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Text1(13).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(16).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub frmLlamPre_DatoSeleccionado(CadenaSeleccion As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaSeleccion <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaSeleccion, 1)
        CadB = Aux
        Aux = Format(ValorDevueltoFormGrid(Text1(7), CadenaSeleccion, 2), FormatoHora)
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaSeleccion, 3)
        CadB = CadB & " AND " & Aux
        Aux = ValorDevueltoFormGrid(Text1(8), CadenaSeleccion, 4)
        CadB = CadB & " AND " & Aux
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmS_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1)
    Text2(1).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

'Private Sub Image1_DblClick()
'Dim L As Long
'Dim C As String
'
'
'        ImprimirImagen
'        Exit Sub
'
'End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte
Dim Observaciones As String

Select Case Index
    Case 4 ' observaciones
        If Modo = 3 Or Modo = 4 Then
            CadenaDesdeOtroForm = Text1(38).Text
        Else
            CadenaDesdeOtroForm = ""
            If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Adodc1.Recordset!observa1, "T")
        End If
        frmFacClienteObser.Modificar = Modo >= 3
        frmFacClienteObser.Text1 = CadenaDesdeOtroForm
        frmFacClienteObser.Show vbModal
        'Llevara DOS VALORES.
        'Si modifica y el texto
        If Modo = 3 Or Modo = 4 Then
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(38).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
        CadenaDesdeOtroForm = ""

    
    Case 5 ' observaciones
        If Modo = 3 Or Modo = 4 Then
            CadenaDesdeOtroForm = Text1(39).Text
        Else
            CadenaDesdeOtroForm = ""
            If Not Adodc1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Adodc1.Recordset!observa2, "T")
        End If
        frmFacClienteObser.Modificar = Modo >= 3
        frmFacClienteObser.Text1 = CadenaDesdeOtroForm
        frmFacClienteObser.Show vbModal
        'Llevara DOS VALORES.
        'Si modifica y el texto
        If Modo = 3 Or Modo = 4 Then
            If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(39).Text = Mid(CadenaDesdeOtroForm, 3)
        End If
        CadenaDesdeOtroForm = ""

    Case 0 'población
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            If Index = 9 Then
                indice = 4
            Else
                PonerFoco Text1(3)
            End If
    Case 1 'clientes
            Set frmCli = New frmFacClientes
            frmCli.DatosADevolverBusqueda = "0|1|"
            frmCli.Show vbModal
            Set frmCli = Nothing
    Case 3 'socios
            Set frmS = New frmGesSocios
            frmS.DatosADevolverBusqueda = "0|1|"
            frmS.Show vbModal
            Set frmS = Nothing
    End Select
End Sub


Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
    Select Case Index
        Case 0
            indice = 1
        Case 1
            indice = 11
        Case 2
            indice = 18
        Case 3
            indice = 19
        Case 4
            indice = 20
        Case 5
            indice = 21
    End Select
    Set frmCal = New frmCal
    If Text1(indice).Text <> "" Then PonerFormatoFecha Text1(indice)
    frmCal.Fecha = Now
    If Text1(indice).Text <> "" Then frmCal.Fecha = CDate(Text1(indice).Text)
    Screen.MousePointer = vbDefault
    frmCal.Show vbModal
    If IsDate(Fecha) Then
        If Fecha <> "0:00:00 " Then Text1(indice) = Fecha
    End If
    Set frmCal = Nothing
    PonerFoco Text1(indice)
End Sub

Private Sub lw1_DblClick()
    ImprimirImagen
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub CargarDocumentos()
    'Hacemos las acciones
    If Modo = 2 Then CargaDatosLW
End Sub


Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim encontrado As String

If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
If Text1(Index).Text <> "" Then Text1(Index) = UCase(Text1(Index).Text)

Select Case Index
    Case 0 'socio
        If Modo = 1 Then Exit Sub
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = Format(Text1(Index).Text, "00000")
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de socio introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                Text2(1).Text = encontrado
            End If
        End If
    Case 8
        If Text1(Index).Text <> "" Then Text1(Index).Text = Format(Text1(Index).Text, "0000")
    Case 13 'cliente
        If Modo = 1 Then Exit Sub
        If Text1(Index).Text <> "" Then
            Text1(Index).Text = Format(Text1(Index).Text, "000000")
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", Text1(Index).Text, "T")
            If encontrado = "" Then
                MsgBox "El código de cliente introducido no existe.", vbExclamation
                PonerFoco Text1(Index)
            Else
                Text1(16).Text = encontrado
            End If
        End If
    Case 3 'CP
        If Text1(Index) <> "" Then
            'Poblacion
            Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, encontrado)
            'provincia
            Text1(Index + 2).Text = encontrado
        End If
    Case 28 To 36
        PonerFormatoDecimal Text1(Index), 6
    Case 7, 17, 22 To 25
        If Text1(Index).Text <> "" Then PonerFormatoHora Text1(Index)
    Case 1, 11, 18 To 21
        If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
        
    Case 43
        PonerFormatoEntero Text1(Index)
End Select
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
            
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
        Case 8  'imprimir
            printNou
    End Select
End Sub
Private Sub mnNuevo_Click()
         BotonAnyadir
End Sub
Private Sub BotonAnyadir()
'Añadir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vacía los TextBox
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    Text1(1).Text = Date
    Check1(0).Value = 0
    Check1(1).Value = 0
    Check1(2).Value = 0
    PonerFoco Text1(1)
End Sub
Private Sub mnModificar_Click()
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub
Private Sub BotonModificar()
'Prepara el Form para Modificar
Dim DeVarios As Boolean
Dim Sql As String
On Error GoTo EModificar

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    '[Monica]04/02/2015: cargamos las variables para el log
    Socio = Text1(0).Text
    Tfno = Text1(12).Text
    clien = Text1(13).Text
    Ident = Text1(9).Text
    Autor = Text1(15).Text
    TipoSer = Combo1.ListIndex
    Licencia = Text1(6).Text
    Matricula = Text1(10).Text
    TxFactu = Check1(0).Value
    TxAbo = Check1(1).Value
    TxVali = Check1(2).Value
    ImpCompra = Text1(29).Text
    ImpVta = Text1(30).Text
    LiqSoc = Check1(3).Value
    FacCli = Check1(4).Value
    Usuario = Text1(14).Text
    Nombre = Text1(16).Text
    Domicilio = Text1(2).Text
    Destino = Text1(42).Text
    numero = Text1(40).Text
    Puerta = Text1(41).Text
    Poblacion = Text1(4).Text
    Obser = Text1(37).Text
    Obs1 = Text1(38).Text
    Obs2 = Text1(39).Text
    
    
    If Combo1.Text = "0" Then Text1(13).Enabled = False
    imgFecha(0).Enabled = False
'    imgBuscar(3).Enabled = False
    If Combo1.Text = "0" Then
        imgBuscar(1).Enabled = False
    Else
        imgBuscar(1).Enabled = True
    End If
    PonerFoco Text1(0)
   
EModificar:
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub mnEliminar_Click()
    BotonEliminar
End Sub
Private Sub mnSalir_Click()
    Unload Me
End Sub
Private Sub BotonEliminar()
Dim msg As String
Dim Sql As String

On Error GoTo EEliminar

msg = "Esta seguro que desea eliminar la llamada del día:" & Text1(1).Text & "?"
If MsgBox(msg, vbYesNo) = vbYes Then
    NumRegElim = Adodc1.Recordset.AbsolutePosition
    Sql = "Delete from shilla where fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    Sql = Sql & "' and codsocio=" & Text1(0).Text & " and numeruve=" & Text1(8).Text
    conn.Execute Sql
End If

If SituarDataTrasEliminar(Adodc1, NumRegElim) Then
    PonerCampos
End If

EEliminar:
If Err.Number <> 0 Then
    MsgBox "Error al eliminar conductor." & Err.Description
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
    cadB1 = ""
End Sub
Private Sub mnBuscar_Click()
    BotonBuscar
End Sub
Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        
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
    CargarDocumentos
    
End Sub

Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Adodc1.RecordSource = CadenaConsulta
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
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
    
    If Combo1.Text = "1" Then
        If Text1(13).Text <> "" Then
            encontrado = DevuelveDesdeBD(conAri, "nomclien", "scliente", "codclien", Text1(13).Text, "T")
            If encontrado = "" Then
                Text1(16).Text = encontrado
            End If
        End If
    End If
    If Text1(0).Text <> "" Then
        encontrado = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", Text1(0).Text, "T")
        If encontrado <> "" Then
            Text2(1).Text = encontrado
        End If
    End If
    
    
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub MandaBusquedaPrevia(CadB As String)
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim Cad As String
'Dim Tabla As String
'Dim Titulo As String
'
'    'Llamamos a al form
'    '##A mano
'    Cad = ""
'    Cad = Cad & ParaGrid(Text1(1), 14, "Fecha")
'    Cad = Cad & ParaGrid(Text1(7), 14, "Hora")
'    Cad = Cad & ParaGrid(Text1(0), 14, "Socio")
'    Cad = Cad & ParaGrid(Text1(8), 14, "Vehiculo")
'
'    Tabla = "shilla"
'    Titulo = "Histórico"
'    If Cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = Cad
'        frmB.vTabla = Tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|2|3|"
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

    
    
    Set frmLlamPre = New frmGesHisLlamPrev

    frmLlamPre.DatosADevolverBusqueda = "0|1|2|4|"
    frmLlamPre.cWhere = CadB
    frmLlamPre.Show vbModal

    Set frmLlamPre = Nothing


End Sub

Private Sub PosicionarData()
Dim Cad As String, Indicador As String
Dim vWhere As String

    If Not Adodc1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Me.Adodc1, vWhere, Indicador) Then
             PonerModo 2
             
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
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
    
    Sql = "(fecha='" & Format(Text1(1).Text, FormatoFecha) & "' and hora='" & Format(Text1(7).Text, FormatoHora)
    Sql = Sql & "' and numeruve=" & Text1(8).Text & ")"
    
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function





Private Sub printNou()
    
Dim indRPT As Byte
Dim cadParam As String
Dim Cad As String
Dim numParam As Byte
Dim nomDocu As String

    indRPT = 56 'Historico de llamadas
    If PonerParamRPT(indRPT, cadParam, numParam, nomDocu, pImprimeDirecto, pPdfRpt) Then
        With frmImprimir2
            .cadTabla2 = "shilla"
            .Informe2 = nomDocu '"rGesHisLlam.rpt"
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
    End If
End Sub


Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Cad = ""
'    If Me.DesdeFichaCliente Then
        '
    Cad = " WHERE fecha=" & DBSet(FechaServ, "F") & " AND hora= " & DBSet(HoraServ, "H") & " AND numeruve=" & DBSet(NumerUve, "N")
        
'    Else
'        'Tengo YA el codigo de la factura
'                '******************************************************
'                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
'                If hcoCodTipoM = "FTI" Then
'                    'no hay albaran directamente va a factura de ticket
'
'                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
'                    cad = "SELECT COUNT(*) FROM scafaccli "
'                    cad = cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
'                    If RegistrosAListar(cad) > 0 Then
'                        cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
'                    Else
'                        cad = ""
'                    End If
'                Else
'                    If hcoCodTipoM = "FAM" Then
'                        cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
'                    End If
'                End If
'                '******************************************************
'
'                If cad = "" Then
'                    'En la smoval estaba e mov. de ALbaran
'                    cad = "SELECT codtipom,numfactu,fecfactu FROM scafaccli1 "
'                    cad = cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
'
'                    Set RS = New ADODB.Recordset
'                    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                    If Not RS.EOF Then 'where para la factura
'                        cad = " WHERE codtipom='" & RS!codtipom & "' AND numfactu= " & RS!NumFactu & " AND fecfactu=" & DBSet(RS!FecFactu, "F")
'                    Else
'                        cad = " WHERE numfactu=-1"
'                    End If
'                    RS.Close
'                    Set RS = Nothing
'                End If
'
'    End If
    ObtenerSelFactura = Cad
End Function

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Marcar validados
            MarcarValidados
    End Select

End Sub

Private Sub MarcarValidados()
Dim Sql As String
Dim NRegs As Long
Dim vMens As String
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim frmMens As frmMensajes
    On Error GoTo eMarcarValidados

    If Me.Adodc1.Recordset.EOF Then Exit Sub

    If CadB <> "" Then
        ContinuarValidar = False
        ComoValidados = False
        
        Set frmMens = New frmMensajes
        frmMens.OpcionMensaje = 32
        frmMens.Show vbModal
        Set frmMens = Nothing

        b = False
        
        If ContinuarValidar Then
            NRegs = Adodc1.Recordset.RecordCount
            vMens = "Va a actualizar " & NRegs & " registros. " & vbCrLf & vbCrLf & " ¿ Desea continuar ?"
            If MsgBox(vMens, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                conn.BeginTrans
                
                b = True
                Adodc1.Recordset.MoveFirst
                I = 1
                Label1(38).visible = True
                While Not Adodc1.Recordset.EOF
                    Label1(38).Caption = "Servicio " & I & " de " & NRegs
                    DoEvents
                    
                    If ComoValidados Then
                        Adodc1.Recordset!validado = 1
                        Adodc1.Recordset.Update
                    Else
                        Adodc1.Recordset!validado = 0
                        Adodc1.Recordset.Update
                    End If
                
                    Adodc1.Recordset.MoveNext
                    I = I + 1
                Wend
                Label1(38).visible = False
            
            
                MsgBox "Proceso realizado correctamente", vbExclamation
                conn.CommitTrans
            
                Adodc1.Recordset.MoveFirst
                PonerCampos

                Screen.MousePointer = vbDefault
            
                Exit Sub
            
            
            End If
        End If
    End If

    Exit Sub
    
eMarcarValidados:
    MuestraError Err.Number, "Marcar Validados", Err.Description
    Label1(38).visible = False
    If b Then conn.RollbackTrans
End Sub



Private Sub ToolbarAux0_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Fuerzo que se vean las lineas
    
    Select Case Button.Index
        Case 1 ' insertar
            InsertarDesdeFichero
            
            CargaDatosLW
            
        Case 3 ' eliminar
            EliminarImagen
        
    End Select
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

'***********************************************************
'Private Sub CargaColumnas()
'Dim Columnas As String
'Dim Ancho As String
'Dim Alinea As String
'Dim Formato As String
'Dim Ncol As Integer
'Dim C As ColumnHeader
'
'    Frame8.visible = False
'
'    ' Documentos
''    Label16.Caption = "Imágenes"
'    Columnas = "Código|Nombre|Documento|Id|Tipo|"
'    Ancho = "1000|6000|0|0|0|"
'    'vwColumnRight =1  left=0   center=2
'    Alinea = "0|0|0|0|0|"
'    'Formatos
'    Formato = "|||||"
'    Ncol = 5
'
'    Frame8.visible = True
'
'
'
''    'Fecha incio busquedas
''    Text1(46).Text = Format(imgFecha(3).Tag, "dd/mm/yyyy")
'    'Guardo la opcion en el tag
'    lw1.Tag = 0 & "|" & Ncol & "|"
'
'    lw1.ColumnHeaders.Clear
'
'    For NumRegElim = 1 To Ncol
'         Set C = lw1.ColumnHeaders.Add()
'         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
'         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
'         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
'         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
'    Next NumRegElim
'End Sub
'
'Private Sub CargaDatosLW()
'Dim C As String
'Dim bs As Byte
'    bs = Screen.MousePointer
'    C = Me.lblIndicador.Caption
'    lblIndicador.Caption = "Leyendo Imágenes"
'    lblIndicador.Refresh
'    CargaDatosLW2
'    Me.lblIndicador.Caption = C
'    Screen.MousePointer = bs
'End Sub
'
'Private Sub CargaDatosLW2()
'Dim Cad As String
'Dim Rs As ADODB.Recordset
'Dim It As ListItem
'Dim ElIcono As Integer
'Dim GroupBy As String
'Dim Orden As String
'Dim C As String
'
'
'    On Error GoTo ECargaDatosLW
'
'    If Modo <> 2 Then Exit Sub
'
''    For NumRegElim = 1 To Toolbar2.Buttons.Count
''        If Toolbar2.Buttons(NumRegElim).Value = tbrPressed Then
''            ElIcono = Toolbar3.Buttons(NumRegElim).Image
''            Exit For
''        End If
''    Next
'
'    'Fecha incio busquedas
''    Text3(0).Text = Format(imgFec(3).Tag, "dd/mm/yyyy")
'
'    If Me.Adodc1.Recordset.EOF Then Exit Sub
'
'    Cad = "select h.orden, h.descripfich, h.campo, h.codigo, h.docum from rfichdocs h WHERE "
'    Cad = Cad & " fecha=" & DBSet(Me.Adodc1.Recordset!Fecha, "F")
'    Cad = Cad & " and hora = " & DBSet(Me.Adodc1.Recordset!hora, "H")
'    Cad = Cad & " and numeruve = " & DBSet(Me.Adodc1.Recordset!NumerUve, "N")
'    GroupBy = ""
'    BuscaChekc = "orden"
'
'
'    'El group by
'    If GroupBy <> "" Then Cad = Cad & " GROUP BY " & GroupBy
'
'    'El ORDER BY
'    If CByte(RecuperaValor(lw1.Tag, 1)) = 1 Then BuscaChekc = Orden
'
'    'BuscaChekc="" si es la opcion de precios especiales
'    If Cad <> "" Then Cad = Cad & " ORDER BY " & BuscaChekc & " DESC"
'    BuscaChekc = ""
'
'    lw1.ListItems.Clear
'
'    CargarArchivos
'
'    Exit Sub
'
'ECargaDatosLW:
'    MuestraError Err.Number
'    Set Rs = Nothing
'End Sub
'
'Private Sub CargarArchivos()
'Dim C As String
'Dim L As Long
'Dim Rs As ADODB.Recordset
'Dim nFile As Long
'
'    ProcesarCarpetaImagenes
'
'    C = "Select * from rfichdocs where fecha=" & DBSet(Text1(1).Text, "F")
'    C = C & " and hora = " & DBSet(Text1(7), "H")
'    C = C & " and numeruve = " & DBSet(Text1(8).Text, "N") & " ORDER BY orden "
''    Me.lblCarga2.Caption = "Leyendo desde BD "
''    Me.lblCarga2.Refresh
'    Adodc2.ConnectionString = conn
'    Adodc2.RecordSource = C
'    Adodc2.Refresh
'
'    If Adodc2.Recordset.EOF Then
'        'NO HAY NINGUNA
'        CargarIMG ""
'    Else
'        'LEEMOS LAS IMAGENES
''        InsertandoImg = True
'        While Not Adodc2.Recordset.EOF
'            L = Adodc2.Recordset!Codigo
''            Me.lblCarga2.Caption = "Leyendo desde BD " & L & "       " & adodc1.Recordset.AbsolutePosition & " de " & adodc1.Recordset.RecordCount
''            lblCarga2.Refresh
'            C = App.Path & "\" & CarpetaIMG & "\" & L
'            If DBLet(Adodc2.Recordset!DOCUM) <> "0" Then
'                C = App.Path & "\" & CarpetaIMG & "\" & Adodc2.Recordset!DOCUM
'            End If
'            If Dir(C) <> "" Then
'                AnyadirAlListview C, True
'            Else
'                If LeerBinary(Adodc2.Recordset!campo, C) Then
'                    AnyadirAlListview C, True
'                End If
'            End If
'            Adodc2.Recordset.MoveNext
'        Wend
'
'
'
''        InsertandoImg = False
'        If lw1.ListItems.Count > 0 Then CargarIMG lw1.ListItems(1).SubItems(2)
'    End If
'
'    Set Adodc2.Recordset = Nothing
'End Sub
'
'Private Sub ProcesarCarpetaImagenes()
'Dim C As String
'Dim MiNombre As String
'
'    On Error GoTo EProcesarCarpetaImagenes
'    C = App.Path & "\" & CarpetaIMG
'    If Dir(C, vbDirectory) = "" Then
'        MkDir C
'    Else
'        On Error Resume Next
'        If Dir(C & "\*.*", vbArchive) <> "" Then 'Kill c & "\*.*"
'            MiNombre = Dir(C & "\*.*")   ' Recupera la primera entrada.
'            Do While MiNombre <> ""   ' Inicia el bucle.
'               ' Ignora el directorio actual y el que lo abarca.
'               If MiNombre <> "." And MiNombre <> ".." Then
'                    Kill C & "\" & MiNombre
'               End If
'               MiNombre = Dir   ' Obtiene siguiente entrada.
'            Loop
'        End If
'        On Error GoTo EProcesarCarpetaImagenes
'
'    End If
'
'    Exit Sub
'EProcesarCarpetaImagenes:
'    MuestraError Err.Number, "ProcesarCarpetaImagenes"
'End Sub
'
'
'
'Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
'Dim J As Integer
'Dim Aux As String
'Dim It As ListItem
'Dim Contador As Integer
'    If Dir(vpaz, vbArchive) = "" Then
'        MsgBox "No existe el archivo: " & vpaz, vbExclamation
'    Else
'        'List1.AddItem vpaz
'        Set It = lw1.ListItems.Add()
''        It.SmallIcon = 23
'
''        If DesdeBD Then
''            J = InStrRev(vpaz, "\") + 1
''            Aux = Mid(vpaz, J)
''            It.Text = "Código " & Aux
''            If Not IsNumeric(Aux) Then It.SmallIcon = 9
''            It.SubItems(2) = Aux
''
''        Else
''            Contador = Contador + 1
'            It.Text = Me.Adodc2.Recordset!Orden '"Nuevo " & Contador
''        End If
'
'        It.SubItems(1) = Me.Adodc2.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
'        It.SubItems(2) = vpaz
'        It.SubItems(3) = Me.Adodc2.Recordset.Fields(0)
'        Set It = Nothing
'    End If
'End Sub
'
'
'Private Function CargarIMG(Archivo As String) As Boolean
'    On Error Resume Next
'    Screen.MousePointer = vbHourglass
''    lblCarga2.Caption = "Cargando ..."
''    lblCarga2.Refresh
'    CargarIMG = False
'
'    If InStr(1, Archivo, ".pdf") <> 0 Then
'        Me.Image1.Picture = LoadPicture(App.Path & "\pdf.dat")
'    Else
'        If InStr(1, Archivo, ".tif") <> 0 Then
'            Me.Image1.Picture = LoadPicture(App.Path & "\tif.dat")
'        Else
'            If InStr(1, Archivo, ".png") Then
'                Me.Image1.Picture = LoadPicture(App.Path & "\png.dat")
'            Else
'                Me.Image1.Picture = LoadPicture(Archivo)
'            End If
'        End If
'    End If
'
'    If Err.Number <> 0 Then
'        MsgBox Err.Description, vbExclamation
'    Else
'        CargarIMG = True
'    End If
''    lblCarga2.Caption = lblCarga2.Tag
'    Screen.MousePointer = vbDefault
'End Function


'***********************************************************
'********************************************************
'
' FUNCIONES CORRESPONDIENTES A LA INSERCION DE DOCUMENTOS
'
'********************************************************
Private Function InsertarDesdeFichero() As Boolean
Dim CADENA As String
Dim Carpeta As String
Dim Aux As String
Dim J As Integer
Dim C As String
Dim Rs As ADODB.Recordset
Dim L As Long


    Fichero = ""
    cd1.FileName = ""
    cd1.InitDir = "c:\"
    cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    cd1.MaxFileSize = 1024 * 30
    cd1.Filter = "Archivos PDF|*.pdf|Archivos Jpg|*.jpg"
    cd1.ShowOpen
    cd1.MaxFileSize = 256
    cd1.CancelError = False
    
    If cd1.FileName = "" Then
        InsertarDesdeFichero = False
        Exit Function
    End If
    
    If FileLen(cd1.FileName) / 1000 > 1024 Then
        MsgBox "No se permite insertar ficheros de tamaño superior a 1 M", vbExclamation
        InsertarDesdeFichero = False
        Exit Function
    End If
    
    
'    '******* Cambiamos cursor
    Screen.MousePointer = vbHourglass

    J = InStr(1, cd1.FileName, Chr(0))
    CADENA = cd1.FileName
    TipoDocu = 0
    If InStr(1, cd1.FileName, "pdf") <> 0 Then TipoDocu = 1
    Fichero = CADENA
        
            
    Screen.MousePointer = vbDefault
    
    txtAux3(4).Text = CCur(DevuelveValor("select max(orden) from shilla_fichdocs where fecha = " & DBSet(Text1(1), "F") & " and hora = " & DBSet(Text1(7).Text, "H") & " and numeruve = " & DBSet(Text1(8).Text, "N")) + 1)
    txtAux3(5).Text = Dir(CADENA)
    
    C = "Select max(codigo) from shilla_fichdocs"
    Set Rs = New ADODB.Recordset
    Rs.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    L = 0
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then L = Rs.Fields(0)
    End If
    L = L + 1
    Rs.Close
    
    ' es nuevo
    C = "insert into shilla_fichdocs (codigo, fecha, hora, numeruve, orden, docum) values"
    C = C & " (" & DBSet(L, "N") & "," & DBSet(Text1(1).Text, "F") & "," & DBSet(Text1(7).Text, "H") & "," & DBSet(Text1(8).Text, "N") & "," & DBSet(txtAux3(4).Text, "N") & "," & DBSet(txtAux3(5).Text, "T") & ")"
    conn.Execute C
    
    Espera 0.2
    
    'Abro parar guardar el binary
    C = "Select * from shilla_fichdocs where codigo =" & L '& " and codsocio = " & DBSet(RecuperaValor(vDatos, 1), "N")
    Adodc2.ConnectionString = conn
    Adodc2.RecordSource = C
    Adodc2.Refresh
'
    If Adodc2.Recordset.EOF Then
        'MAAAAAAAAAAAAL

    Else
        'Guardar
        GuardarBinary Adodc2.Recordset!campo, Fichero
        Adodc2.Recordset.Update
    End If
    
End Function



Private Sub CargaDatosLW()
Dim C As String
Dim bs As Byte
    bs = Screen.MousePointer
    C = Me.lblIndicador.Caption
    lblIndicador.Caption = "Leyendo "
    lblIndicador.Refresh
    CargaDatosLW2
    Me.lblIndicador.Caption = C
    Screen.MousePointer = bs
End Sub

Private Sub CargaDatosLW2()
Dim Cad As String
Dim Rs As ADODB.Recordset
Dim It As ListItem
Dim ElIcono As Integer
Dim GroupBy As String
Dim Orden As String
Dim C As String


    On Error GoTo ECargaDatosLW
    
    If Modo <> 2 Then Exit Sub
    
    If Adodc1.Recordset.EOF Then Exit Sub
    
    Cad = "select h.orden, h.campo, h.codigo, h.docum from shilla_fichdocs h WHERE "
    Cad = Cad & " fecha=" & DBSet(Adodc1.Recordset!Fecha, "F")
    Cad = Cad & " and hora=" & DBSet(Adodc1.Recordset!hora, "H")
    Cad = Cad & " and numeruve=" & Adodc1.Recordset!NumerUve
    GroupBy = ""
    BuscaChekc = "orden"
        
    
    
    'BuscaChekc="" si es la opcion de precios especiales
    Cad = Cad & " ORDER BY 1"
    
    lw1.ListItems.Clear
    
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    While Not Rs.EOF
        Set It = lw1.ListItems.Add

        It.Text = Rs!Orden '"Nuevo " & Contador

        It.SubItems(1) = Rs!DOCUM  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        It.SubItems(2) = App.Path & "\" & CarpetaIMG & "\" & Rs!DOCUM
        It.SubItems(3) = Rs!Codigo

        Set It = Nothing

        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    ProcesarCarpetaImagenes
    
    
    Exit Sub
ECargaDatosLW:
    MuestraError Err.Number
    Set Rs = Nothing
    
End Sub

Private Sub CargarArchivos()
Dim C As String
Dim L As Long


    ProcesarCarpetaImagenes


    C = "Select * from shilla_fichdocs where fecha=" & DBSet(Text1(1).Text, "F")
    C = C & " and hora = " & DBSet(Text1(7).Text, "H")
    C = C & " and numeruve = " & DBSet(Text1(8).Text, "N")
    C = C & " and codigo = " & DBSet(lw1.SelectedItem.SubItems(3), "N")

    Adodc2.ConnectionString = conn
    Adodc2.RecordSource = C
    Adodc2.Refresh

    If Adodc2.Recordset.EOF Then
        'NO HAY NINGUNA
    Else
        'LEEMOS LA IMAGEN
        L = Adodc2.Recordset!Codigo
        C = App.Path & "\" & CarpetaIMG & "\" & L
        If DBLet(Adodc2.Recordset!DOCUM) <> "0" Then
            C = App.Path & "\" & CarpetaIMG & "\" & Adodc2.Recordset!DOCUM
        End If
        LeerBinary Adodc2.Recordset!campo, C
    End If

End Sub



Private Sub ProcesarCarpetaImagenes()
Dim C As String
Dim MiNombre As String

    On Error GoTo EProcesarCarpetaImagenes
    
    C = App.Path & "\" & CarpetaIMG
    If Dir(C, vbDirectory) = "" Then
        MkDir C
    Else
        On Error Resume Next
        If Dir(C & "\*.*", vbArchive) <> "" Then 'Kill c & "\*.*"
            MiNombre = Dir(C & "\*.*")   ' Recupera la primera entrada.
            Do While MiNombre <> ""   ' Inicia el bucle.
               ' Ignora el directorio actual y el que lo abarca.
               If MiNombre <> "." And MiNombre <> ".." Then
                    Kill C & "\" & MiNombre
               End If
               MiNombre = Dir   ' Obtiene siguiente entrada.
            Loop
        End If
        On Error GoTo EProcesarCarpetaImagenes
    
    End If
    
    Exit Sub
EProcesarCarpetaImagenes:
    MuestraError Err.Number, "ProcesarCarpetaImagenes"
End Sub

Private Sub AnyadirAlListview(vpaz As String, DesdeBD As Boolean)
Dim J As Integer
Dim Aux As String
Dim It As ListItem
Dim Contador As Integer
    If Dir(vpaz, vbArchive) = "" Then
'        MsgBox "No existe el archivo: " & vpaz, vbExclamation
    Else
        Set It = lw1.ListItems.Add()

        It.Text = Me.Adodc2.Recordset!Orden '"Nuevo " & Contador
        
        It.SubItems(1) = Me.Adodc2.Recordset.Fields(5)  'Abs(DesdeBD)   'DesdeBD 0:NO  numero: el codigo en la BD
        It.SubItems(2) = vpaz
        It.SubItems(3) = Me.Adodc2.Recordset.Fields(0)
        
        Set It = Nothing
    End If
End Sub


'Private Sub ImprimirImagen()
'Dim NFic As Long
'Dim vAdobe As String
'
'   CargarArchivos
'
'   Call ShellExecute(Me.hwnd, "Open", Me.lw1.SelectedItem.SubItems(2), "", "", 1)
'
'End Sub


Private Sub EliminarImagen()
Dim Sql As String
Dim Mens As String
    
    On Error GoTo eEliminarImagen

    Mens = "Va a proceder a eliminar de la lista correspondiente a la llamada. " & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf & vbCrLf
    
    If MsgBox(Mens, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "delete from shilla_fichdocs where fecha = " & DBSet(Text1(1).Text, "F") & " and hora = " & DBSet(Text1(7).Text, "H") & " and numeruve = " & DBSet(Text1(8).Text, "N") & " and codigo = " & Me.lw1.SelectedItem.SubItems(3)
        conn.Execute Sql
        FicheroAEliminar = lw1.SelectedItem.SubItems(2)
        CargaDatosLW
        
    End If
    Exit Sub

eEliminarImagen:
    MuestraError Err.Number, "Eliminar imágen", Err.Description
End Sub


Private Sub CargarColumnas()
Dim Columnas As String
Dim Ancho As String
Dim Alinea As String
Dim Formato As String
Dim Ncol As Integer
Dim C As ColumnHeader

    Columnas = "Código|Nombre|Documento|Id|Tipo|"
    Ancho = "1000|5350|0|0|0|"
    'vwColumnRight =1  left=0   center=2
    Alinea = "0|0|0|0|0|"
    'Formatos
    Formato = "|||||"
    Ncol = 5

    lw1.Tag = "5|" & Ncol & "|"
    
    lw1.ColumnHeaders.Clear
    
    For NumRegElim = 1 To Ncol
         Set C = lw1.ColumnHeaders.Add()
         C.Text = RecuperaValor(Columnas, CInt(NumRegElim))
         C.Width = RecuperaValor(Ancho, CInt(NumRegElim))
         C.Alignment = Val(RecuperaValor(Alinea, CInt(NumRegElim)))
         C.Tag = RecuperaValor(Formato, CInt(NumRegElim))
    Next NumRegElim


End Sub

