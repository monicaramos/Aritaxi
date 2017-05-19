VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPubliHcoFacSoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Facturas Publicidad Socios"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   400
      Width           =   12015
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
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   390
         Width           =   1305
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   7560
         MaxLength       =   40
         TabIndex        =   5
         Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwww aq"
         Top             =   330
         Width           =   4290
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
         Left            =   6720
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Cod. Socio|N|N|0|999999|sfactusoc|codsocio|000000|S|"
         Text            =   "Text1"
         Top             =   330
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   1
         Tag             =   "Tipo Factura|T|N|||sfactusoc|codtipom||S|"
         Text            =   "Text3"
         Top             =   405
         Width           =   735
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
         Left            =   2670
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||sfactusoc|fecfactu|dd/mm/yyyy|S|"
         Top             =   405
         Width           =   1240
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
         Tag             =   "Nº Factura|N|N|||sfactusoc|numfactu|0000000|S|"
         Text            =   "Text1 7"
         Top             =   405
         Width           =   980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Contabilizado"
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
         Left            =   4050
         TabIndex        =   3
         Tag             =   "Contabilizado|N|N|0|1|sfactusoc|intconta||N|"
         Top             =   285
         Width           =   1695
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
         Index           =   0
         Left            =   5820
         TabIndex        =   42
         Top             =   330
         Width           =   555
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   6450
         ToolTipText     =   "Buscar socio"
         Top             =   330
         Width           =   240
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
         Index           =   29
         Left            =   2670
         TabIndex        =   41
         Top             =   120
         Width           =   1245
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
         Index           =   28
         Left            =   180
         TabIndex        =   40
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Fact."
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
         Left            =   1320
         TabIndex        =   39
         Top             =   120
         Width           =   1125
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9240
      Top             =   1680
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
      Height          =   4170
      Left            =   120
      TabIndex        =   19
      Tag             =   "Fecha Oferta|F|N|||scapre|fecentre|dd/mm/yyyy|N|"
      Top             =   1245
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   7355
      _Version        =   393216
      Style           =   1
      Tabs            =   1
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
      TabPicture(0)   =   "frmPubliHcoFacSoc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameCliente"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameFactura"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame FrameFactura 
         Height          =   1980
         Left            =   180
         TabIndex        =   28
         Top             =   1920
         Width           =   11685
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
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
            Left            =   7320
            MaxLength       =   15
            TabIndex        =   53
            Tag             =   "Total Factura|N|N|||sfactusoc|totalfac|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1515
         End
         Begin VB.TextBox Text1 
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
            Index           =   18
            Left            =   5100
            MaxLength       =   15
            TabIndex        =   52
            Tag             =   "Importe IVA 1|N|N|||sfactusoc|impoiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1515
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
            Left            =   4380
            MaxLength       =   5
            TabIndex        =   51
            Tag             =   "% IVA 1|N|S|0|99.90|sfactusoc|porciva1|#0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   555
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
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   50
            Tag             =   "Base Imponible 1|N|N|||sfactusoc|baseiva1|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   1515
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
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   49
            Tag             =   "Cod. IVA 1|N|S|0|9999|sfactusoc|codiiva1|0000|N|"
            Text            =   "Text1 7"
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox Text1 
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
            Index           =   14
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   48
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1515
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
            Left            =   120
            MaxLength       =   15
            TabIndex        =   47
            Tag             =   "Imp.Bruto|N|N|||sfactusoc|importel|#,###,###,##0.00|N|"
            Text            =   "Text1 7"
            Top             =   360
            Width           =   1515
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
            Left            =   2160
            TabIndex        =   46
            Tag             =   "Concepto|T|N|||sfactusoc|concepto||N|"
            Text            =   "Text1 7"
            Top             =   1530
            Width           =   6675
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1800
            ToolTipText     =   "Ver observaciones"
            Top             =   1560
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Concepto"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   44
            Top             =   1560
            Width           =   1605
         End
         Begin VB.Label Label1 
            Caption         =   "Importe IVA"
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
            Index           =   37
            Left            =   5340
            TabIndex        =   43
            Top             =   720
            Width           =   1215
         End
         Begin VB.Line Line1 
            X1              =   2040
            X2              =   2040
            Y1              =   960
            Y2              =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "Desglose IVA"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   42
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   1605
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
            Left            =   4350
            TabIndex        =   36
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "TOTAL FACTURA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Index           =   39
            Left            =   7320
            TabIndex        =   35
            Top             =   720
            Width           =   1530
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
            Index           =   38
            Left            =   6840
            TabIndex        =   34
            Top             =   960
            Width           =   135
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
            TabIndex        =   33
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label1 
            Caption         =   "Base imponible"
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
            Left            =   2790
            TabIndex        =   32
            Top             =   720
            Width           =   1485
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
            Left            =   1920
            TabIndex        =   31
            Top             =   360
            Width           =   135
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
            Index           =   14
            Left            =   2160
            TabIndex        =   30
            Top             =   120
            Width           =   1725
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
            Index           =   10
            Left            =   240
            TabIndex        =   29
            Top             =   120
            Width           =   1485
         End
      End
      Begin VB.Frame FrameCliente 
         Caption         =   "Datos Socio"
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
         Height          =   1455
         Left            =   180
         TabIndex        =   20
         Top             =   360
         Width           =   11685
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
            Index           =   12
            Left            =   6840
            MaxLength       =   3
            TabIndex        =   45
            Tag             =   "Forma de Pago|N|N|0|999|sfactusoc|codforpa|000|N|"
            Text            =   "Text1"
            Top             =   645
            Width           =   540
         End
         Begin VB.TextBox Text1 
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
            Height          =   315
            Index           =   11
            Left            =   6840
            MaxLength       =   30
            TabIndex        =   11
            Text            =   "Text1 Text1 Text1 Text1 Text22"
            Top             =   990
            Width           =   2445
         End
         Begin VB.TextBox Text1 
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
            Index           =   9
            Left            =   1425
            MaxLength       =   6
            TabIndex        =   9
            Text            =   "Text15"
            Top             =   990
            Width           =   630
         End
         Begin VB.TextBox Text1 
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
            Index           =   10
            Left            =   2055
            MaxLength       =   30
            TabIndex        =   10
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwaq"
            Top             =   990
            Width           =   3405
         End
         Begin VB.TextBox Text1 
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
            Index           =   7
            Left            =   3495
            MaxLength       =   20
            TabIndex        =   7
            Text            =   "12345678911234567899"
            Top             =   285
            Width           =   1965
         End
         Begin VB.TextBox Text1 
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
            Index           =   6
            Left            =   1425
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "123456789"
            Top             =   285
            Width           =   1110
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
            Index           =   15
            Left            =   7470
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   22
            Text            =   "Text2"
            Top             =   645
            Width           =   3405
         End
         Begin VB.TextBox Text1 
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
            Index           =   8
            Left            =   1425
            MaxLength       =   35
            TabIndex        =   8
            Text            =   "Text1 wwwwwwwwwwwwwwwwwwwwwwwwww aq"
            Top             =   645
            Width           =   4030
         End
         Begin VB.Image imgBuscar 
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1155
            ToolTipText     =   "Buscar población"
            Top             =   1005
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
            Left            =   5760
            TabIndex        =   27
            Top             =   990
            Width           =   735
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
            TabIndex        =   26
            Top             =   990
            Width           =   975
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
            Left            =   2625
            TabIndex        =   25
            Top             =   285
            Width           =   885
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
            TabIndex        =   24
            Top             =   285
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Forma Pago"
            Height          =   255
            Index           =   15
            Left            =   5760
            TabIndex        =   23
            Top             =   645
            Width           =   855
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   5
            Left            =   6600
            ToolTipText     =   "Buscar forma de pago"
            Top             =   645
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
            TabIndex        =   21
            Top             =   645
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   5460
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
         TabIndex        =   16
         Top             =   120
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
      Left            =   11070
      TabIndex        =   13
      Top             =   5520
      Width           =   1035
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
      Left            =   9780
      TabIndex        =   12
      Top             =   5520
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Rectifica"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "0"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Factura"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   18
         Top             =   120
         Width           =   1215
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
      Left            =   11070
      TabIndex        =   14
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
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
      Begin VB.Menu mnRectifica 
         Caption         =   "&Rectifica"
         Enabled         =   0   'False
         Shortcut        =   ^R
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLineas 
         Caption         =   "&Lineas"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^L
         Visible         =   0   'False
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImprimirAlbaran 
         Caption         =   "Imprimir &albarán"
         Enabled         =   0   'False
         Shortcut        =   ^A
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
Attribute VB_Name = "frmPubliHcoFacSoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento
Public hcoCodSocio As String ' codigo de socio

Public publicidad As Boolean
Public DesdeFichaSocio As Boolean


Dim PrimeraVez As Boolean
Dim NombreTabla As String
Dim Ordenacion As String
Dim CadenaConsulta As String
Private kCampo As Integer
Private btnPrimero As Byte
Private HaDevueltoDatos As Boolean
Private Modo As Byte
Private BuscaChekc As String
Private CodTipoMov As String

Private cadFormula As String
Private cadParam As String
Private numParam As Byte

Private WithEvents frmC As frmGesSocios
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmFP As frmFacFormasPago
Attribute frmFP.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal
Attribute frmCP.VB_VarHelpID = -1

Dim cadB1 As String

Dim UnaVez As Boolean

Private Sub ComprobarDatosTotales()
Dim i As Byte

    For i = 13 To 14
        Text1(i).Text = ComprobarCero(Text1(i).Text)
    Next i
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
    ComprobarDatosTotales
    
    'comprobamos datos OK de la tabla scafac
    b = CompForm(Me, 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not b Then Exit Function
    
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

'Private Function ModificaAlbxFac() As Boolean
'Dim SQL As String
'Dim B As Boolean
'On Error GoTo EModificaAlb
'
'    ModificaAlbxFac = False
'    If Data1.Recordset.EOF Then Exit Function
'
'    'comprobar datos OK de la scafac1
'     B = CompForm(Me, 2) 'Comprobar formato datos ok de la cabecera alb: opcion=2
'    If Not B Then Exit Function
'
'    SQL = "UPDATE scafpa SET codtrab2=" & DBSet(Text3(0).Text, "N", "S") & ", "
'    SQL = SQL & "codtrab1=" & DBSet(Text3(1).Text, "N", "S")
'    If Me.FrameObserva.visible Then
'        SQL = SQL & ", observa1=" & DBSet(Text3(4).Text, "T")
'        SQL = SQL & ", observa2=" & DBSet(Text3(5).Text, "T")
'        SQL = SQL & ", observa3=" & DBSet(Text3(6).Text, "T")
'        SQL = SQL & ", observa4=" & DBSet(Text3(7).Text, "T")
'        SQL = SQL & ", observa5=" & DBSet(Text3(8).Text, "T")
'    End If
'    SQL = SQL & ObtenerWhereCP(True)
'    SQL = SQL & " AND numalbar=" & Data1.Recordset.Fields!NumAlbar
'    conn.Execute SQL
'    ModificaAlbxFac = True
'
'EModificaAlb:
'If Err.Number <> 0 Then MuestraError Err.Number, "Modificar Albaranes de factura", Err.Description
'End Function


Private Function ModificarFactura(Optional sqlLineas As String) As Boolean
'si se ha modificado la linea de slifac, añadir a la transaccion la modificación de la linea y recalcular
Dim bol As Boolean
Dim MenError As String
Dim Sql As String
Dim vFactu As CFacturaSoc
On Error GoTo EModFact

    bol = False
    conn.BeginTrans
    ConnConta.BeginTrans
    
    If sqlLineas <> "" Then
        'actualizar el importe de la linea modificada
        MenError = "Modificando lineas de Factura."
        conn.Execute sqlLineas
    End If
    
'    'recalcular las bases imponibles x IVA
'    MenError = "Recalcular importes IVA"
'    bol = ActualizarDatosFactura
    bol = True
    If bol Then
        'modificamos la scafpc
        MenError = "Modificando cabecera de factura"
        bol = ModificaDesdeFormulario(Me, 1)
        
'        If bol Then
'            'Si es proveedor de varios actualizar datos proveedor en tabla:sprvar
'            MenError = "Modificando datos proveedor varios"
'            bol = ActualizarProveVarios(Text1(2).Text, Text1(4).Text)
'        End If
        
        If bol Then
'            MenError = "Modificando albaranes de factura"
'            'modificar la tabla: scafpa
'            bol = ModificaAlbxFac
            
            If bol Then 'si se ha modificado la factura
                MenError = "Actualizando en Tesoreria"
                'y eliminar de tesoreria conta.spagop los registros de la factura
                
                'antes de Eliminar en las tablas de la Contabilidad
                Set vFactu = New CFacturaSoc
                bol = vFactu.LeerDatos3(Text1(4).Text, Text1(0).Text, Text1(2).Text)
                
                If bol Then
                    'Eliminar de la spagop
                    If vParamAplic.ContabilidadNueva Then
                        Sql = " codmacta='" & vFactu.CtaSocio & "' AND numfactu=" & Data1.Recordset.Fields!NumFactu & ""
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from pagos WHERE " & Sql
                    Else
                        Sql = " ctaprove='" & vFactu.CtaSocio & "' AND numfactu=" & Data1.Recordset.Fields!NumFactu & ""
                        Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                        ConnConta.Execute "Delete from spagop WHERE " & Sql
                    End If
                    'Volvemos a grabar en TESORERIA. Tabla de Contabilidad: sconta.spagop
                    If bol Then
                        bol = vFactu.InsertarEnTesoreria(MenError)
                    End If
                End If
                Set vFactu = Nothing
            End If
        End If
    End If

EModFact:
     If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        bol = False
    End If
    If bol Then
        conn.CommitTrans
        ConnConta.CommitTrans
        ModificarFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        ModificarFactura = False
        MenError = "Actualizando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
    End If
End Function

Private Function CalcularDatosFactura() As Boolean
Dim i As Integer
Dim vFactu As CFacturaCom
Dim FacOK As Boolean

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For i = 22 To 38
         Text1(i).Text = ""
    Next i
    
    
    Set vFactu = New CFacturaCom
    vFactu.DtoPPago = CCur(ComprobarCero(Text1(16).Text))
    vFactu.DtoGnral = CCur(ComprobarCero(Text1(17).Text))
    vFactu.Proveedor = Text1(4).Text
    
    If vFactu.CalcularDatosFactura(ObtenerWhereCP(False), NombreTabla, "") Then
        FacOK = True
        Text1(22).Text = vFactu.BrutoFac
        Text1(23).Text = vFactu.ImpPPago
        Text1(24).Text = vFactu.ImpGnral
        Text1(25).Text = vFactu.BaseImp
        Text1(26).Text = QuitarCero(vFactu.TipoIVA1)
        Text1(27).Text = QuitarCero(vFactu.TipoIVA2)
        Text1(28).Text = QuitarCero(vFactu.TipoIVA3)
        Text1(29).Text = vFactu.PorceIVA1
        Text1(30).Text = vFactu.PorceIVA2
        Text1(31).Text = vFactu.PorceIVA3
        Text1(32).Text = vFactu.BaseIVA1
        Text1(33).Text = vFactu.BaseIVA2
        Text1(34).Text = vFactu.BaseIVA3
        Text1(35).Text = vFactu.ImpIVA1
        Text1(36).Text = vFactu.ImpIVA2
        Text1(37).Text = vFactu.ImpIVA3
        Text1(38).Text = vFactu.TotalFac
        FormatoDatosTotales
    Else
        FacOK = False
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
    Set vFactu = Nothing
    CalcularDatosFactura = FacOK
End Function


Private Function ActualizarDatosFactura() As Boolean
Dim vFactu As CFacturaCom
Dim cadSel As String

    Set vFactu = New CFacturaCom
    cadSel = ObtenerWhereCP(False)
    cadSel = "slifpc." & cadSel
    vFactu.DtoPPago = CCur(Text1(11).Text)
    vFactu.DtoGnral = CCur(Text1(12).Text)
    
'    'Si tiene RETENCION
'    If Me.FrmRetencionSocios.visible Then
'        vFactu.PorRet = ImporteFormateado(Text1(32).Text)
'        vFactu.ImpRet2 = ImporteFormateado(Text1(33).Text)
'    End If

    
    
    If vFactu.CalcularDatosFactura(cadSel, "scafpa", "slifpc") Then
        Text1(14).Text = vFactu.BrutoFac
        Text1(15).Text = vFactu.ImpPPago
        Text1(16).Text = vFactu.ImpGnral
        Text1(17).Text = vFactu.BaseImp
        Text1(18).Text = vFactu.TipoIVA1
        Text1(19).Text = vFactu.TipoIVA2
        Text1(20).Text = vFactu.TipoIVA3
        Text1(21).Text = vFactu.PorceIVA1
        Text1(22).Text = vFactu.PorceIVA2
        Text1(23).Text = vFactu.PorceIVA3
        Text1(24).Text = vFactu.BaseIVA1
        Text1(25).Text = vFactu.BaseIVA2
        Text1(26).Text = vFactu.BaseIVA3
        Text1(27).Text = vFactu.ImpIVA1
        Text1(28).Text = vFactu.ImpIVA2
        Text1(29).Text = vFactu.ImpIVA3
        Text1(30).Text = vFactu.TotalFac
'        If Me.FrmRetencionSocios.visible Then
'            Text1(32).Text = vFactu.PorRet
'            Text1(33).Text = vFactu.ImpRet2
'        End If
        
        FormatoDatosTotales
        
        ActualizarDatosFactura = True
    Else
        ActualizarDatosFactura = False
        MuestraError Err.Number, "Recalculando Factura", Err.Description
    End If
    Set vFactu = Nothing
End Function


Private Sub cmdAceptar_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'INSERTAR
            If DatosOk Then
                If InsertarCabecera Then
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    PonerModo 2
                End If
            End If




        Case 4  'MODIFICAR
            If DatosOk Then
               If ModificarFactura Then
               
                                        
                    Set LOG = New cLOG
                    LOG.Insertar 8, vUsu, "Factura modificada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf
                    Set LOG = Nothing
               
               
                    Espera 0.2
                    TerminaBloquear
                    PosicionarData
'                    FormatoDatosTotales
'                    i = Data1.Recordset.AbsolutePosition
'                    PonerCamposLineas
'                    SituarDataPosicion Data1, CLng(i), ""
                End If
            End If
            
'         Case 5 'InsertarModificar LINEAS
'            If ModificaLineas = 2 Then 'MODIFICAR lineas
'                If ModificarLinea Then
'
'                        'INSERTA LOG
'                        '-------------------------------------------------
'                        Set LOG = New cLOG
'                        BuscaChekc = "     Alb : " & Data2.Recordset!NumAlbar & "   Linea: " & Data2.Recordset!numlinea
'                        BuscaChekc = "Modificar linea: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & BuscaChekc
'                        LOG.Insertar 8, vUsu, BuscaChekc
'                        Set LOG = Nothing
'                        BuscaChekc = ""
'
'                    TerminaBloquear
'                    CargaGrid DataGrid1, Data2, True
'                    ModificaLineas = 0
'                    PonerBotonCabecera True
'                    BloquearTxt Text2(16), True
'
'                    LLamaLineas Modo, 0, "DataGrid1"
'                    PosicionarData
'                Else
'                    TerminaBloquear
'                End If
'                Me.DataGrid1.Enabled = True
'            End If
    End Select
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarCabecera() As Boolean
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String
Dim NumFactu As Long
Dim vSocio As CSocio
Dim bol As Boolean
Dim devuelve As Long
Dim Existe As Boolean
Dim MenError As String
Dim vFacSoc As CFacturaSoc
Dim CtaBanco As String


'    On Error GoTo EInsertarCab
'
'    CodTipoMov = "FRQ"
'
'    bol = False
'
'    conn.BeginTrans
'    ConnConta.BeginTrans
'
'    Set vSocio = New CSocio
'    If vSocio.LeerDatos(Text1(4).Text) Then
'        NumFactu = vSocio.ConseguirContador(CodTipoMov)
'        If NumFactu = -1 Then bol = False
'        Do
'            NumFactu = vSocio.ConseguirContador(CodTipoMov)
'            SQL = "select numfactu from rfactusoc where codtipom = " & DBSet(CodTipoMov, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(Text1(2).Text, "F") & " and codsocio = " & DBSet(vSocio.Codigo, "N")
'            devuelve = DevuelveValor(SQL) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
'            If devuelve <> 0 Then
'                'Ya existe el contador incrementarlo
'                Existe = True
'                vSocio.IncrementarContador (CodTipoMov)
'                NumFactu = vSocio.ConseguirContador(CodTipoMov)
'            Else
'                Existe = False
'            End If
'        Loop Until Not Existe
'        Text1(0).Text = NumFactu
'
'        SQL = CadenaInsertarDesdeForm(Me)
'        If SQL <> "" Then
'
'            MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
'            conn.Execute SQL, , adCmdText
'
'            Set vFacSoc = New CFacturaSoc
'            '[Monica]22/11/2013: iban
'            vFacSoc.CCC_Iban = vSocio.Iban
'            vFacSoc.CCC_Entidad = vSocio.Banco
'            vFacSoc.CCC_Oficina = vSocio.Sucursal
'            vFacSoc.CCC_CC = vSocio.DigControl
'            vFacSoc.CCC_CTa = vSocio.CuentaBan
'            vFacSoc.ForPago = Text1(12).Text
'            vFacSoc.tipoMov = CodTipoMov
'            vFacSoc.NumFactu = Text1(0).Text
'            vFacSoc.FecFactu = Text1(2).Text
'            '[Monica]10/07/2012: Tiene que estar en negativo en la spagop
'            vFacSoc.TotalFac = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(19).Text))) '* (-1)
'            vFacSoc.ImpRet2 = Text1(20).Text
'            vFacSoc.Socio = Text1(4).Text
'
'            vFacSoc.CtaSocio = vSocio.CtaSocioLiq
'
'            cadban = ""
'
'            Set frmBan = New frmFacBancosPropios
'            frmBan.DatosADevolverBusqueda = "1|"
'            frmBan.Show vbModal
'            Set frmBan = Nothing
'
'            CtaBanco = cadban ' InputBox("Introduzca el Banco de pago: ", "Tesoreria", , 5000, 4000)
'
'            If CtaBanco = "" Then
'                MsgBox "No ha seleccionado cuenta de banco.", vbExclamation
'                bol = False
'            Else
'                bol = True
'                vFacSoc.CuentaPrev = DevuelveDesdeBDNew(conAri, "sbanpr", "codmacta", "codbanpr", CtaBanco, "N")
'            End If
'
'            If bol Then bol = ActualizarRetencion(vSocio.UveSocio, vFacSoc, True)
'
'            '[Monica]10/07/2012: Tiene que estar en negativo en la spagop
'            If bol Then bol = vFacSoc.InsertarEnTesoreria(MenError)   'vFacSoc.InsertarEnTesoreriaCobro("", MenError)
'
'            If bol Then bol = vSocio.IncrementarContador(CodTipoMov)
'
'            Set vFacSoc = Nothing
'
'        End If
'
'        Text1(0).Text = Format(Text1(0).Text, "0000000")
'    End If
'    Set vSocio = Nothing
'
'EInsertarCab:
'    Screen.MousePointer = vbDefault
'
'    If Err.Number <> 0 Then
'        MenError = "Insertando Factura." & vbCrLf & "----------------------------" & vbCrLf & MenError
'        MuestraError Err.Number, MenError, Err.Description
'        bol = False
'    End If
'    If bol Then
'        conn.CommitTrans
'        ConnConta.CommitTrans
'        InsertarCabecera = True
'    Else
'        conn.RollbackTrans
'        ConnConta.RollbackTrans
'        InsertarCabecera = False
'    End If
End Function



Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
'             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub FormatoDatosTotales()
Dim i As Byte

    For i = 16 To 19
'        Text1(I).Text = QuitarCero(Text1(I).Text)
        FormateaCampo Text1(i)
    Next i
    
'    For i = 24 To 26
'        If Text1(i).Text <> "" Then
'            'Si la Base Imp. es 0
'            If CSng(Text1(i).Text) = 0 Then
'                Text1(i).Text = QuitarCero(Text1(i).Text)
'                Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
'                Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
'                Text1(i + 3).Text = QuitarCero(Text1(i + 3).Text)
'            Else
'                FormateaCampo Text1(i)
'                FormateaCampo Text1(i - 3)
'                FormateaCampo Text1(i - 6)
'                FormateaCampo Text1(i + 3)
'            End If
'        Else 'No hay Base Imponible
'            Text1(i - 3).Text = QuitarCero(Text1(i - 3).Text)
'            Text1(i - 6).Text = QuitarCero(Text1(i - 6).Text)
'            Text1(i + 3).Text = ""
'        End If
'    Next i
'
'    If Me.FrmRetencionSocios.visible Then
'        FormateaCampo Text1(32)
'        FormateaCampo Text1(33)
'    End If
End Sub

Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
'            LimpiarDataGrids
            PonerModo 0
'            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
'        Case 5 'Lineas Detalle
'            TerminaBloquear
'            BloquearTxt Text2(16), True
'            If ModificaLineas = 1 Then 'INSERTAR
'                ModificaLineas = 0
'                DataGrid1.AllowAddNew = False
'                If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
'            End If
'            DataGrid2.Enabled = True
'            ModificaLineas = 0
'            LLamaLineas Modo, 0, "DataGrid1"
'            PonerBotonCabecera True
'            Me.DataGrid1.Enabled = True
    End Select

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus()
    Text1(1).Text = Mid(Combo1.List(Combo1.ListIndex), 1, 3)
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If UnaVez Then
        UnaVez = False
        If hcoCodMovim <> "" Then
            If Data1.Recordset.EOF Then
                PonerCadenaBusqueda
            Else
                PonerCampos
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    UnaVez = True
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
     'Icono de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscar(2).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscar(3).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscar(5).Picture = frmPpal.imgIcoForms.ListImages(1).Picture

    ' ICONITOS DE LA BARRA
    btnPrimero = 12
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Botón Buscar
        .Buttons(2).Image = 2   'Botón Todos
        
        .Buttons(4).Image = 3   'Insertar Nuevo
        .Buttons(5).Image = 4   'Modificar
        .Buttons(6).Image = 5   'Borrar
        .Buttons(7).Image = 16 'Imprimir
        .Buttons(10).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Último
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
        
    '## A mano
    NombreTabla = "sfactusoc"
    Ordenacion = " ORDER BY sfactusoc.codtipom, sfactusoc.numfactu, sfactusoc.fecfactu "
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Dim T1 As Single
    T1 = Timer
    
    CargaCombo
    
    Combo1.ListIndex = 0
    
    cadB1 = "sfactusoc.codtipom in ('FPS','FRQ')"
    
    CadenaConsulta = "Select * from " & NombreTabla ' & " where codtipom is null and " & cadB1
    
    If hcoCodMovim <> "" Then
        'Se llama desde Dobleclick en frmAlmMovimArticulos
        'como tenemos aqui el nº de albaran, buscar a que factura corresponde
        'en la scafaccli1
        CadenaConsulta = CadenaConsulta & ObtenerSelFactura & " and " & cadB1
'        CadenaConsulta = CadenaConsulta & " WHERE codtipom='" & hcoCodTipoM & "' AND numalbar= " & hcoCodMovim
    Else
        'CadenaConsulta = CadenaConsulta & " where numfactu=-1"
        'Cambio sugerido por Msoler
        'mariela 01/07/2010
        CadenaConsulta = CadenaConsulta & " WHERE codtipom is null and numfactu is null and fecfactu is null and " & cadB1
    End If

    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    
    Data1.Refresh
    
   
    If hcoCodMovim = "" Then
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            BotonBuscar
        End If
'        CargaGrid DataGrid1, Data2, False
        'Poner los grid sin apuntar a nada
'        LimpiarDataGrids
        PrimeraVez = False
    Else
        If Data1.Recordset.EOF Then
            PonerModo 0
        Else
            PonerModo 2
        End If
    End If
                
End Sub

Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
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

Private Sub HacerBusqueda()
Dim CadB As String
    
    If Text1(4).Text <> "" Then Text1(5).Text = ""
    
    CadB = ObtenerBusqueda(Me, False, BuscaChekc)
    
    If CadB = "" Then
        CadB = cadB1
    Else
        CadB = CadB & " and " & cadB1
    End If
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
'        cadB = cadB & " and scafac.codtipom='" & CodTipoMov & "'" 'Solo seleccionamos los del Movimiento, aqui los FAV
        CadenaConsulta = "select sfactusoc.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY codtipom,codsocio,numfactu,fecfactu " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    Me.Check1.Value = 0
    Me.Combo1.ListIndex = -1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    hcoCodMovim = ""
    hcoCodTipoM = "" 'Codigo detalle de Movimiento(ALC)
    hcoFechaMov = "" 'fecha del movimiento
    hcoCodSocio = "" ' codigo de socio
    DesdeFichaSocio = False
End Sub

Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Formulario Mantenimiento C. Postales
Dim indice As Byte
Dim devuelve As String

    indice = 9
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'CPostal
    Text1(indice + 1).Text = ObtenerPoblacion(Text1(indice).Text, devuelve)  'Poblacion
    'provincia
    Text1(indice + 2).Text = devuelve

End Sub

Private Sub frmFP_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Formas de Pago
Dim indice As Byte
    indice = 12
    Text1(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Forma Pago
    Text2(indice + 3).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Forma Pago
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

'    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Cliente
            PonerFoco Text1(4)
            Set frmC = New frmGesSocios
            frmC.DatosADevolverBusqueda = "0"
            frmC.Show vbModal
            Set frmC = Nothing
            PonerFoco Text1(4)
            PonerDatosCliente
        Case 5 'forma de pago
            indice = 12
            PonerFoco Text1(indice)
            Set frmFP = New frmFacFormasPago
            frmFP.DatosADevolverBusqueda = "0"
            frmFP.Show vbModal
            PonerFoco Text1(12)
            Set frmFP = Nothing
        Case 2 'codpobla
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 9
            PonerFoco Text1(indice)
            
        Case 3 ' observaciones
            If Modo = 5 Or Modo = 0 Then
            Else
                
                If Modo = 3 Or Modo = 4 Then
                    CadenaDesdeOtroForm = Text1(3).Text
                Else
                    CadenaDesdeOtroForm = ""
                    If Not Me.Data1.Recordset.EOF Then CadenaDesdeOtroForm = DBLet(Data1.Recordset!Concepto, "T")
                End If
                frmFacClienteObser.Modificar = Modo >= 3
                frmFacClienteObser.Text1 = CadenaDesdeOtroForm
                frmFacClienteObser.Show vbModal
                'Llevara DOS VALORES.
                'Si modifica y el texto
                If Modo = 3 Or Modo = 4 Then
                    If RecuperaValor(CadenaDesdeOtroForm, 1) = "1" Then Text1(3).Text = Mid(CadenaDesdeOtroForm, 3)
                End If
                CadenaDesdeOtroForm = ""
            End If
        
    End Select

End Sub

Private Sub PonerDatosCliente()
Dim Cad As String

    If Text1(4).Text = "" Then Exit Sub


    Set miRsAux = New ADODB.Recordset
    
    Cad = "select * from sclien where codclien=" & Text1(4).Text
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(5).Text = miRsAux!nomclien
        Text1(6).Text = miRsAux!nifClien
        Text1(7).Text = miRsAux!telclie1
        Text1(8).Text = miRsAux!domclien
        Text1(9).Text = miRsAux!codpobla
        Text1(10).Text = miRsAux!pobclien
        Text1(11).Text = miRsAux!proclien
    End If
    
End Sub


Private Sub mnBuscar_Click()
    BotonBuscar
End Sub



Private Sub mnModificar_Click()

    If vUsu.Nivel > 0 Then
        MsgBox "No tiene permiso para realizar la accion", vbExclamation
        Exit Sub
    End If
    'bloquea la tabla cabecera de factura: scafac
    If BLOQUEADesdeFormulario(Me) Then
        'bloquear la tabla cabecera de albaranes de la factura: sfactusoc
        BotonModificar
    End If
    
End Sub

Private Sub BotonModificar()
Dim DeVarios As Boolean
Dim EnTesoreria  As String
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada3(EnTesoreria) Then
        TerminaBloquear
        Exit Sub
    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFocoChk Me.Check1
        
'    'Inserto en slog
'
'    Set LOG = New cLOG
'    If EnTesoreria <> "" Then EnTesoreria = "Tesoreria: " & vbCrLf & EnTesoreria
'    EnTesoreria = Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EnTesoreria
'    EnTesoreria = "Pulsa mod factura: " & EnTesoreria
'    LOG.Insertar 8, vUsu, EnTesoreria
'    Set LOG = Nothing
'    Espera 0.3
'    '
    
End Sub

Private Function FactContabilizada(ByRef EstaEnTesoreria As String) As Boolean
Dim cta As String, numasien As String
    
    On Error GoTo EContab
    
    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        
        cta = vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(4).Text, "0000")
'        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        If cta <> "" Then
            numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
            If numasien <> "" Then
                FactContabilizada = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
            Else
                FactContabilizada = False
            End If
        Else
            FactContabilizada = True
            Exit Function
        End If
    Else
        FactContabilizada = False
    End If
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function

Private Function FactContabilizada3(ByRef EstaEnTesoreria As String) As Boolean
Dim cta As String, numasien As String
    
    On Error GoTo EContab
    
    'comprabar que se puede modificar/eliminar la factura
    If Me.Check1.Value = 1 Then 'si esta contabilizada
        'comprobar en la contabilidad si esta contabilizada
        
'        Cta = vParamAplic.Raiz_Cta_Soc_publi & Format(Text1(4).Text, "0000")
''        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
'        If Cta <> "" Then
'            numasien = DevuelveDesdeBDNew(conConta, "cabfactprov", "numasien", "codmacta", Cta, "T", , "numfacpr", Text1(0).Text, "T", "fecfacpr", Text1(1).Text, "F")
'            If numasien <> "" Then
                FactContabilizada3 = True
                MsgBox "La factura esta contabilizada y no se puede modificar ni eliminar.", vbInformation
                Exit Function
'            Else
'                FactContabilizada = False
'            End If
'        Else
'            FactContabilizada = True
'            Exit Function
'        End If
    Else
        FactContabilizada3 = False
    End If
    
EContab:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Factura contabilizada", Err.Description
End Function



Private Sub mnRectifica_Click()
    If Modo = 5 Then 'Añadir lineas
'         BotonAnyadirLinea
    Else 'Añadir Cabecera
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If

End Sub

Private Sub BotonAnyadir()
'Añadir registro en tabla de cabecera de Pedidos: scaped (Cabecera)
Dim NomTraba As String
Dim Cad As String
Dim RS As ADODB.Recordset

    LimpiarCampos 'Vacía los TextBox
    'Poner los grid sin apuntar a nada
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    NomTraba = ""

    Text1(2).Text = Format(Now, "dd/mm/yyyy") 'Fecha Albaran
    Text1(1).Text = "FRQ"
    Text1(15).Text = vParamAplic.IVA_REA
    Text1(17).Text = DevuelveDesdeBDNew(conConta, "tiposiva", "porceiva", "codigiva", vParamAplic.IVA_REA, "N")
    PonerFoco Text1(2)
End Sub


Private Sub Text1_LostFocus(Index As Integer)
Dim devuelve As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 4 'socio
            If Modo = 1 Then 'Modo=1 Busqueda
                '-- Laura 12/01/2007
                'Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, NombreTabla, "nomclien")
                Text1(5).Text = PonerNombreDeCod(Text1(Index), conAri, "sclien", "nomclien", "codclien", Text1(Index).Text, "N")
                '--
            Else
                PonerDatosSocio
            End If
            
            
        Case 12 'forma de pago
            If Text1(Index).Text <> "" Then
                devuelve = DevuelveDesdeBD(conAri, "nomforpa", "sforpa", "codforpa", Text1(Index).Text, "T")
                If devuelve = "" Then
                    MsgBox "El código de forma de pago introducido no existe.", vbExclamation
                    PonerFoco Text1(Index)
                    Exit Sub
                Else
                    Text2(Index + 3).Text = devuelve
                End If
            End If
        
        Case 13, 16, 18, 19
            PonerFormatoDecimal Text1(Index), 3
            
       Case 17
            PonerFormatoDecimal Text1(Index), 7
              
        
    End Select
End Sub

Private Sub PonerDatosSocio()
Dim Cad As String

    If Text1(4).Text = "" Then Exit Sub

    Set miRsAux = New ADODB.Recordset
    
    Cad = "select * from sclien where codclien=" & Text1(4).Text
    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        Text1(5).Text = miRsAux!nomclien
        Text1(6).Text = miRsAux!nifClien
        Text1(7).Text = DBLet(miRsAux!telclie1, "T")
        Text1(8).Text = miRsAux!domclien
        Text1(9).Text = miRsAux!codpobla
        Text1(10).Text = miRsAux!pobclien
        Text1(11).Text = miRsAux!proclien
    End If
    
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: mnBuscar_Click  'Buscar
        Case 2: BotonVerTodos  'Todos


        Case 4: mnRectifica_Click
        Case 5: mnModificar_Click  'Modificar
        Case 6: mnEliminar_Click  'Borrar
        
        Case 7: mnImprimir_Click 'Imprimir Albaran
        
        Case 10: mnSalir_Click    'Salir
            
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub
Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub
Private Sub mnImprimir_Click()
    HacerImpresionFacturas
End Sub
Private Sub HacerImpresionFacturas()
    cadFormula = "({sfactusoc.codtipom}='FPS' and {sfactusoc.numfactu}=" & Text1(0).Text & " and {sfactusoc.codsocio}=" & Text1(4).Text & " and {sfactusoc.fecfactu}= Date(""" & Text1(2).Text & """))"
    LlamarImprimir True
End Sub
Private Sub LlamarImprimir(duplicado As Boolean)
Dim devuelve As String

    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    cadParam = cadParam & "pDuplicado= " & Abs(duplicado) & " |"
    numParam = 2
    
    devuelve = DevuelveDesdeBD(conAri, "letraser", "stipom", "codtipom", Text1(1).Text, "T")
    
    
    With frmImprimir
        'Nuevo. Febrero 2010
        .outClaveNombreArchiv = devuelve & Format(Text1(0).Text, "000")
        .outCodigoCliProv = Text1(4).Text
        .outTipoDocumento = 100
        
        
        .Titulo = "Impresión de Facturas de publicidad"
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .NombreRPT = DevuelveDesdeBDNew(conAri, "scryst", "documrpt", "codcryst", "48", "N")
    '------ > Listado 47 = rFacPubli.rpt
        .Opcion = 101
        .ConSubInforme = False
        .Show vbModal
    End With

End Sub
Private Sub BotonVerTodos()

    cadB1 = "sfactusoc.codtipom in ('FPS','FRQ') "

    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia "" & cadB1
    Else
        lblIndicador.Caption = "Preparando bus."
        lblIndicador.Refresh
        LimpiarCampos
        DoEvents
        
        CadenaConsulta = "Select sfactusoc.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & cadB1

        lblIndicador.Caption = "Obteniendo reg."
        lblIndicador.Refresh
        PonerCadenaBusqueda
    End If
End Sub
Private Sub mnEliminar_Click()
    BotonEliminar

End Sub

Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (sfactusoc)
Dim Cad As String
Dim EstaEnTesoreria As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'solo se puede modificar la factura si no esta contabilizada
    If FactContabilizada3(EstaEnTesoreria) Then Exit Sub
    
    Cad = "Cabecera de Facturas." & vbCrLf
    Cad = Cad & "-------------------------------------      " & vbCrLf & vbCrLf
    Cad = Cad & "Va a eliminar la Factura:            "
    Cad = Cad & vbCrLf & "Tipo:  " & Text1(1).Text
    Cad = Cad & vbCrLf & "Nº Fact.:  " & Format(Text1(0).Text, "0000000")
    Cad = Cad & vbCrLf & "Fecha:  " & Format(Text1(2).Text, "dd/mm/yyyy")

    Cad = Cad & vbCrLf & vbCrLf & " ¿Desea Eliminarla? "

    'Borramos
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        CodTipoMov = Text1(1).Text
        
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
        
            
            Set LOG = New cLOG
            LOG.Insertar 8, vUsu, "Factura eliminada: " & Text1(1).Text & Text1(0).Text & " " & Text1(2).Text & vbCrLf & EstaEnTesoreria
            Set LOG = Nothing
        
            If SituarDataTrasEliminar(Data1, NumRegElim) Then
                PonerCampos
            Else
                LimpiarCampos
                'Poner los grid sin apuntar a nada
                PonerModo 0
            End If
        End If
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Albaran", Err.Description
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
            PonerFoco Text1(kCampo)
        End If
        lblIndicador.Caption = ""
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

Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo

    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
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
    If Modo = 2 Then
        If Not Data1.Recordset.EOF Then
            If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
        End If
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
          
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    
    Me.Check1.Enabled = (Modo = 1)
    
    b = (Modo <> 1)
    'Campos Nº Factura bloqueado y en azul
    BloquearTxt Text1(0), b, True
    BloquearTxt Text1(3), b And Modo <> 4 'referencia
    
    'Importes siempre bloqueados, excepto para busquedas. ivas y aportacion tb bloqueado
    For i = 13 To 19
        BloquearTxt Text1(i), (Modo <> 1) And (Modo <> 3)
    Next i
    'Campo B.Imp y Imp. IVA siempre en azul
    BloquearTxt Text1(14), True
    Text1(14).BackColor = &HFFFFC0
    BloquearTxt Text1(18), True
    Text1(18).BackColor = &HFFFFC0
    
    If Modo <> 1 Then
        Text1(18).BackColor = &HFFFFC0
    End If
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
'    BloquearImg imgBuscar(2), True
'    BloquearImg imgBuscar(0), (Modo = 2 Or Modo = 0 Or Modo = 5)
'    BloquearImg imgBuscar(5), (Modo = 2 Or Modo = 0 Or Modo = 5)
'    BloquearImg imgBuscar(3), (Modo = 0)
    
    BloquearImg imgBuscar(2), True
    BloquearImg imgBuscar(0), (Modo <> 1) '(Modo = 2 Or Modo = 0 Or Modo = 5 Or Modo = 4)
    BloquearImg imgBuscar(5), (Modo <> 1) And (Modo <> 4)  '(Modo = 2 Or Modo = 0 Or Modo = 5)
    
    Me.Combo1.visible = (Modo = 1)
    
    
'    For i = 0 To 5
'        If i <> 1 And i <> 4 Then
'            Me.imgBuscar(i).Enabled = B
'        End If
'    Next i
    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub
Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

        b = (Modo = 2) Or (Modo = 5)
        
        'Rectifica
        Toolbar1.Buttons(4).Enabled = b Or Modo = 0
        Me.mnRectifica.Enabled = b Or Modo = 0
        
        'Modificar
        Toolbar1.Buttons(5).Enabled = b
        Me.mnModificar.Enabled = b
        'eliminar
        Toolbar1.Buttons(6).Enabled = (Modo = 2)
        Me.mnEliminar.Enabled = (Modo = 2)
            
        b = (Modo = 2)
        'Imprimir
        Toolbar1.Buttons(7).Enabled = b
        Me.mnImprimir.Enabled = b
        
        b = (Modo >= 3) Or Modo = 1
        'Buscar
        Toolbar1.Buttons(1).Enabled = Not b
        Me.mnBuscar.Enabled = Not b
        'Ver Todos
        Toolbar1.Buttons(2).Enabled = Not b
        Me.mnVerTodos.Enabled = Not b
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub
Private Sub PonerCampos()
Dim BrutoFac As Single
    
    On Error Resume Next
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    BrutoFac = CSng(Text1(13).Text)
    Text1(14).Text = Format(BrutoFac, FormatoImporte)
    
    
    'poner descripcion campos
    Modo = 4
    Text1_LostFocus (12) 'forma de pago
    Modo = 2
    
    'Datos del socio
    PonerDatosCliente
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index = 1 And Modo = 1 Then
        SendKeys "{tab}"
        Exit Sub
    End If
    ConseguirFoco Text1(Index), Modo
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
    Cad = Cad & ParaGrid(Text1(1), 10, "Tipo Fac.")
    Cad = Cad & ParaGrid(Text1(0), 15, "Nº Factura")
    Cad = Cad & ParaGrid(Text1(2), 15, "Fecha Fac.")
    Cad = Cad & ParaGrid(Text1(4), 10, "Socio")
    Cad = Cad & ParaGrid(Text1(5), 50, "Nombre Socio")
    Tabla = NombreTabla
        
    Titulo = "Facturas"
    devuelve = "0|1|2|"
           
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vselElem = 0
        frmB.vConexionGrid = conAri  'Conexión a BD: Aritaxi
        frmB.Show vbModal
        Set frmB = Nothing
        PonerCadenaBusqueda
        Text1(0).Text = Format(Text1(0).Text, "0000000")
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Socios
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)  'Cod Socio
End Sub
Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(2), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    End If
    Screen.MousePointer = vbDefault
End Sub
'En vTesoreria pondremos como estaba el recibo
'Es decir. El  msgbox que pondra al final lo guardo en esta variable
Private Function ComprobarCobroArimoney(vTesoreria As String, LEtra As String, Codfaccl As Long, Fecha As Date) As Boolean
Dim vR As ADODB.Recordset
Dim Cad As String


On Error GoTo EComprobarCobroArimoney
    ComprobarCobroArimoney = False
    Set vR = New ADODB.Recordset
    
    If vParamAplic.ContabilidadNueva Then
        Cad = "Select * from cobros where numserie='" & LEtra & "'"
        Cad = Cad & " AND numfactu =" & Codfaccl
        Cad = Cad & " AND fecfactu =" & DBSet(Fecha, "F")
    Else
        Cad = "Select * from scobro where numserie='" & LEtra & "'"
        Cad = Cad & " AND codfaccl =" & Codfaccl
        Cad = Cad & " AND fecfaccl =" & DBSet(Fecha, "F")
    End If
    '
    vTesoreria = ""
    vR.Open Cad, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If vR.EOF Then
        vTesoreria = "NO se ha encotrado ningun vencimiento en la tesoreria"
    Else
        While Not vR.EOF
            Cad = ""
            If DBLet(vR!codrem, "T") <> "" Then
                Cad = "El cobro asociado a la factura esta remesado(" & vR!codrem & ")"
            Else
                If DBLet(vR!recedocu, "N") = 1 Then
                    Cad = "Documento recibido"
                Else
                
                    If vParamAplic.ContabilidadNueva Then
                            If DBLet(vR!transfer, "N") = 1 Then
                                Cad = "Esta en una transferencia"
                            Else
                               If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                            
                                
                                        'Si hubeira que poner mas coas iria aqui
                            End If 'transfer
                    
                    Else
                
                        If DBLet(vR!Estacaja, "N") = 1 Then
                            Cad = "Cobrado por caja"
                        Else
                            If DBLet(vR!transfer, "N") = 1 Then
                                Cad = "Esta en una transferencia"
                            Else
                               If DBLet(vR!impcobro, "N") > 0 Then Cad = "Esta parcialmente cobrado: " & vR!impcobro
                            
                                
                                        'Si hubeira que poner mas coas iria aqui
                            End If 'transfer
                        End If 'estacaja
                    End If
                End If 'recdedocu
            End If 'remesado
            If Cad <> "" Then vTesoreria = vTesoreria & "Vto: " & vR!numorden & "      " & Cad & vbCrLf
            vR.MoveNext
        Wend
    End If
    vR.Close
    
    
    
    If vTesoreria <> "" Then
        Cad = vTesoreria & vbCrLf & vbCrLf
        If vUsu.Nivel > 1 Then
            MsgBox Cad, vbExclamation
        Else
            Cad = Cad & "¿Continuar?"
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then ComprobarCobroArimoney = True
        End If
    Else
        ComprobarCobroArimoney = True
    End If
            
EComprobarCobroArimoney:
    If Err.Number <> 0 Then MuestraError Err.Number
    Set vR = Nothing
End Function


Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim vFactu As CFacturaSoc
Dim bol As Boolean

    On Error GoTo FinEliminar

'    B = False
'    If Data1.Recordset.EOF Then Exit Function
'
'    conn.BeginTrans
'
'    'Eliminar en las tablas de la Contabilidad
'    '------------------------------------------
'    Letra = ObtenerLetraSerie(Data1.Recordset!codtipom)
'
'    If Letra <> "" Then
'        SQL = " numserie='" & Letra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
'        SQL = SQL & " AND anofaccl=" & Year(Data1.Recordset.Fields!FecFactu)
'
'        'Lineas
'        ConnConta.Execute "Delete from linfact WHERE " & SQL
'
'        'cabecera
'        ConnConta.Execute "Delete from cabfact WHERE " & SQL
'
'        'cobros
'        SQL = " numserie='" & Letra & "' AND codfaccl=" & Data1.Recordset.Fields!NumFactu
'        SQL = SQL & " AND fecfaccl='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
'        ConnConta.Execute "Delete from scobro WHERE " & SQL
'        B = True
'    Else
'        B = False
'    End If
'
'    'Eliminar en tablas de factura de Aritaxi
'    '------------------------------------------
'    If B Then
'        SQL = " " & ObtenerWhereCP(True)
'
'
'        'Eliminar los vencimientos
'        conn.Execute "Delete from svenci " & SQL
'
'        'Cabecera de facturas (sfactusoc)
'        conn.Execute "Delete from " & NombreTabla & SQL
'
'        'Decrementar contador si borramos la ult. factura
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador Data1.Recordset!codtipom, Val(Text1(0).Text)
'        Set vTipoMov = Nothing
'    End If
'
'    B = True

        b = False
        Eliminar = False
        If Data1.Recordset.EOF Then Exit Function
        
        conn.BeginTrans
        ConnConta.BeginTrans
        
        'Eliminar en la tabla pagos de la Contabilidad: spagop
        '------------------------------------------------
'        Cta = DevuelveDesdeBDNew(conAri, "sprove", "codmacta", "codprove", Text1(2).Text, "N")
        
        'antes de Eliminar en las tablas de la Contabilidad
        Set vFactu = New CFacturaSoc
        bol = vFactu.LeerDatos3(Text1(4).Text, Text1(0).Text, Text1(2).Text)
        If bol Then
            
            If vParamAplic.ContabilidadNueva Then
                Sql = " codmacta='" & vFactu.CtaSocio & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                ConnConta.Execute "Delete from pagos WHERE " & Sql
            
            Else
                Sql = " ctaprove='" & vFactu.CtaSocio & "' AND numfactu='" & Data1.Recordset.Fields!NumFactu & "'"
                Sql = Sql & " AND fecfactu='" & Format(Data1.Recordset.Fields!FecFactu, FormatoFecha) & "'"
                ConnConta.Execute "Delete from spagop WHERE " & Sql
            End If
            b = True
            
            
            'Eliminar en tablas de factura de Aritaxi: scafpc, scafpa, slifpc
            '---------------------------------------------------------------
            If b Then
                Sql = " " & ObtenerWhereCP(True)
            
                'Cabecera de facturas (sfactusoc)
                conn.Execute "Delete from " & NombreTabla & Sql
            End If
            
            'Eliminar los movimientos generados por el albaran que genero la factura
            '-----------------------------------------------------------------------
            If b Then
            
            End If
            
    '        b = True
    
        End If
    
        Set vFactu = Nothing
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Factura", Err.Description
        b = False
    End If
    If Not b Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        ConnConta.CommitTrans
        Eliminar = True
    End If
End Function
Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = " codtipom= '" & Text1(1).Text & "' and numfactu= " & Val(Text1(0).Text) & " and fecfactu='" & Format(Text1(2).Text, FormatoFecha) & "' "
    Sql = Sql & " and codsocio = " & Text1(4).Text
    
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function

Private Function ObtenerSelFactura() As String
Dim Cad As String
Dim RS As ADODB.Recordset

    On Error Resume Next

    Cad = ""
    If Me.DesdeFichaSocio Then
        '
        Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F") & " and codsocio = " & DBSet(hcoCodSocio, "N")
        
    Else
        'Tengo YA el codigo de la factura
                '******************************************************
                'laura: esto se puede comentar, ya no hay movimiento FTI en la smoval
                If hcoCodTipoM = "FTI" Then
                    'no hay albaran directamente va a factura de ticket
                    
                    'ver si lo encontramos como factura: codtipom, numfactu,fecfactu
                    Cad = "SELECT COUNT(*) FROM sfactusoc "
                    Cad = Cad & " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    If RegistrosAListar(Cad) > 0 Then
                        Cad = " WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    Else
                        Cad = ""
                    End If
                Else
                    If hcoCodTipoM = "FAM" Then
                        Cad = "  WHERE codtipom='" & hcoCodTipoM & "' AND numfactu= " & hcoCodMovim & " AND fecfactu=" & DBSet(hcoFechaMov, "F")
                    End If
                End If
                '******************************************************
                    
                If Cad = "" Then
                    'En la smoval estaba e mov. de ALbaran
                    Cad = "SELECT codtipom,numfactu,fecfactu FROM sfactusoc "
                    Cad = Cad & " WHERE codtipoa=" & DBSet(hcoCodTipoM, "T") & " AND numalbar=" & hcoCodMovim & " AND fechaalb=" & DBSet(hcoFechaMov, "F")
                    
                    Set RS = New ADODB.Recordset
                    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If Not RS.EOF Then 'where para la factura
                        Cad = " WHERE codtipom='" & RS!codtipom & "' AND numfactu= " & RS!NumFactu & " AND fecfactu=" & DBSet(RS!FecFactu, "F")
                    Else
                        Cad = " WHERE numfactu=-1"
                    End If
                    RS.Close
                    Set RS = Nothing
                End If
    
    End If
    ObtenerSelFactura = Cad
End Function

Private Sub CargaCombo()
Dim RS As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    Combo1.Clear
    
    Sql = "SELECT codtipom,nomtipom FROM stipom WHERE codtipom in  ('FPS')" ','FRQ')"
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Sql = RS!nomtipom
        Sql = Replace(Sql, "Factura", "")
        Combo1.AddItem RS!codtipom & "-" & Sql
        Combo1.ItemData(Combo1.NewIndex) = i
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub

