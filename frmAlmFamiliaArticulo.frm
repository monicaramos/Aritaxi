VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlmFamiliaArticulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Familias de Art�culos"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9990
   Icon            =   "frmAlmFamiliaArticulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Digitos 1er nivel|N|N|||empresa|numdigi1|||"
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3330
      TabIndex        =   36
      Top             =   0
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   37
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
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   3075
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   35
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
      Left            =   7290
      TabIndex        =   33
      Top             =   150
      Width           =   1605
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
      Left            =   2025
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   4890
      Width           =   3585
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
      Left            =   1335
      MaxLength       =   4
      TabIndex        =   9
      Tag             =   "Centro de coste|T|S|||sfamia|codccost||N|"
      Top             =   4890
      Width           =   630
   End
   Begin VB.Frame Frame3 
      Caption         =   "Compras "
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
      Height          =   1095
      Left            =   90
      TabIndex        =   26
      Top             =   3630
      Width           =   9585
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
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Cta. Contable compras|T|N|||sfamia|ctacompr||N|"
         Text            =   "0000000000"
         Top             =   240
         Width           =   1350
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
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   8
         Tag             =   "Cta.Abono Compras|T|N|||sfamia|abocompr||N|"
         Text            =   "Text1"
         Top             =   675
         Width           =   1350
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
         Left            =   4470
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   675
         Width           =   4800
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
         Left            =   4470
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   240
         Width           =   4800
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Abono Compras"
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
         TabIndex        =   30
         Top             =   675
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable Compras"
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
         Left            =   240
         TabIndex        =   29
         Top             =   270
         Width           =   2295
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   1
         Left            =   2700
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   315
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   3
         Left            =   2700
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   720
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ventas "
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
      Height          =   1935
      Left            =   90
      TabIndex        =   17
      Top             =   1560
      Width           =   9585
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
         Left            =   3090
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Cta. Contable Ventas|T|N|||sfamia|ctaventa||N|"
         Text            =   "Text1"
         Top             =   210
         Width           =   1350
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
         Left            =   3090
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Cta. Abono Ventas|T|N|||sfamia|aboventa||N|"
         Text            =   "Text1"
         Top             =   645
         Width           =   1350
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   645
         Width           =   4800
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   210
         Width           =   4800
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
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   1050
         Width           =   4800
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
         Index           =   7
         Left            =   4530
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   1455
         Width           =   4800
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
         Left            =   3090
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Cta. Alternativa Abonos|T|N|||sfamia|abovent1||N|"
         Text            =   "Text1"
         Top             =   1455
         Width           =   1350
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
         Left            =   3090
         MaxLength       =   10
         TabIndex        =   5
         Tag             =   "Cta. Alternativa Ventas|T|N|||sfamia|ctavent1||N|"
         Text            =   "Text1"
         Top             =   1050
         Width           =   1350
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Abono Ventas"
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
         Left            =   240
         TabIndex        =   25
         Top             =   690
         Width           =   2145
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Contable Ventas"
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
         Left            =   240
         TabIndex        =   24
         Top             =   270
         Width           =   2205
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   0
         Left            =   2760
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   270
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   2
         Left            =   2760
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   675
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   5
         Left            =   2760
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Image imgCuentas 
         Height          =   240
         Index           =   4
         Left            =   2760
         Tag             =   "-1"
         ToolTipText     =   "Buscar cuenta contable"
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Alternativa Ventas"
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
         TabIndex        =   23
         Top             =   1110
         Width           =   2565
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Alternativa Abonos"
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
         TabIndex        =   22
         Top             =   1515
         Width           =   2535
      End
   End
   Begin VB.CheckBox chkInstalac 
      Caption         =   "�Es instalaci�n?"
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
      Left            =   7320
      TabIndex        =   2
      Tag             =   "�Es instalaci�n?|N|N|||sfamia|instalac||N|"
      Top             =   1050
      Width           =   2115
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
      Left            =   8520
      TabIndex        =   11
      Top             =   5340
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
      Left            =   3180
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "Denominaci�n familia de Art�culo|T|N|||sfamia|nomfamia||N|"
      Text            =   "Text1"
      Top             =   990
      Width           =   4005
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
      Left            =   870
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "C�digo familia de art�culo|N|N|0|9999|sfamia|codfamia|0000|S|"
      Text            =   "Text"
      Top             =   990
      Width           =   630
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   150
      TabIndex        =   13
      Top             =   5400
      Width           =   2655
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
         Left            =   120
         TabIndex        =   14
         Top             =   150
         Width           =   2355
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
      Left            =   8520
      TabIndex        =   12
      Top             =   5340
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
      Left            =   7230
      TabIndex        =   10
      Top             =   5340
      Width           =   1135
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   480
      Top             =   5430
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
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   0
      Left            =   1005
      ToolTipText     =   "Buscar centro coste"
      Top             =   4890
      Width           =   240
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
      Index           =   4
      Left            =   210
      TabIndex        =   32
      Top             =   4890
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   1710
      TabIndex        =   16
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "C�digo"
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
      TabIndex        =   15
      Top             =   1020
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
Attribute VB_Name = "frmAlmFamiliaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private Const IdPrograma = 205

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmFam As frmBasico2
Attribute frmFam.VB_VarHelpID = -1
Private WithEvents frmCtas As frmBasico2 ' cuentas contables
Attribute frmCtas.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'-------------------------------------------------------------------------
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
Private ModoAnterior As Byte

Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim indCodigo As Integer

Private Sub chkInstalac_GotFocus()
    ConseguirfocoChk Modo
End Sub

Private Sub chkInstalac_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptar_Click()
    
    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    TratarCtaContable
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
    End Select
        
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1 'Busqueda
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
    End Select
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos
    'A�adiremos el boton de aceptar y demas objetos para insertar
    ModoAnterior = Modo 'Para el bot�n cancelar en Modo Insertar
    PonerModo 3
    Text1(0).Text = SugerirCodigoSiguienteStr("sfamia", "codfamia")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue 'vbYellow
    Else 'Modo=1 Busqueda
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
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim cad As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    'El registro de codigo 0 no se puede Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    
    
    '### a mano
    cad = "�Seguro que desea eliminar la Familia de Art�culo?:" & vbCrLf
    cad = cad & vbCrLf & "Cod. : " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Desc.: " & Data1.Recordset.Fields(1)

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        Data1.Recordset.Delete
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        Data1.Recordset.CancelUpdate
        MuestraError Err.Number, "Eliminar Familia de Articulo", Err.Description
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim Im
    'Icono del formulario
    Me.Icon = frmppal.Icon
    If vParamAplic.Descriptores Then Me.Caption = "Categorias Art."
    
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
    
    For Each Im In Me.imgBuscar
        Im.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    For Each Im In Me.imgCuentas
        Im.Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    
    
    
    LimpiarCampos   'Limpia los campos TextBox
    
    'Pone el Tag del primer bot�n de busqueda de cuentas a -1
    'Si tag =-1 abre busqueda en la tabla: sfamia, BD: Aritaxi
    'Si tag>0 abre busqueda en la tabla: Cuentas, BD: Conta
    imgCuentas(0).Tag = "-1"
    Me.imgBuscar(0).Tag = "-1"
        
  
    '## A mano
    NombreTabla = "sfamia"
    Ordenacion = " ORDER BY codfamia"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
       
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " WHERE codfamia=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        BotonBuscar
    End If
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
    Me.chkInstalac.Value = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
Dim indice As Byte
    
    If CadenaDevuelta <> "" Then
        If Val(imgCuentas(0).Tag) >= 0 Then
            'Se llama desde un bot�n de busqueda de Cuentas
            'Recuperar solo el campo c�digo y Descripci�n
            HaDevueltoDatos = True
            Screen.MousePointer = vbHourglass
    
            indice = Val(Me.imgCuentas(0).Tag)
            Text1(indice + 2).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(indice + 2).Text = RecuperaValor(CadenaDevuelta, 2)
            
        ElseIf Val(imgBuscar(0).Tag) >= 0 Then
            'Centro de coste
            Text1(8).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(8).Text = RecuperaValor(CadenaDevuelta, 2)
        Else
            'Recupera todo el registro de Banco Propio
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


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    If CadenaSeleccion <> "" Then
        Text1(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
        Text2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
    End If
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "codfamia = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

End Sub

Private Sub imgBuscar_Click(Index As Integer)
    ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
    
    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
    
    Select Case Index
        Case 0 'Centros de coste de la conta
            Screen.MousePointer = vbHourglass
            Me.imgBuscar(0).Tag = Index
            Set frmB = New frmBuscaGrid
            frmB.vCampos = "Codigo|cabccost|codccost|T||20�Descripci�n|cabccost|nomccost|T||70�"
            frmB.vTabla = "cabccost"
            frmB.vSQL = ""
            HaDevueltoDatos = False
            '###A mano
            frmB.vDevuelve = "0|1|"
            frmB.vTitulo = "Centros de coste"
            frmB.vselElem = 0
            frmB.vConexionGrid = conConta
            
            frmB.Show vbModal
            Set frmB = Nothing
            imgBuscar(0).Tag = -1
            Screen.MousePointer = vbDefault
            PonerFoco Text1(8)
    End Select
End Sub


Private Sub imgCuentas_Click(Index As Integer)

    If Modo = 2 Or Modo = 5 Or Modo = 0 Then Exit Sub
 
    Screen.MousePointer = vbHourglass

    indCodigo = Index + 2

    Set frmCtas = New frmBasico2
    AyudaCuentasContables frmCtas, Text1(indCodigo).Text, "apudirec='S'"
    Set frmCtas = Nothing

    PonerFoco Text1(indCodigo)

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
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not (Index = 0 And KeyCode = 38) Then KEYdown KeyCode
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
        Case 0 'Codigo familia
'            If Text1(Index).Text <> "" Then
             If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod de familia en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
        '#### lo hemos puesto en el evento VALIDATE
'         Case 2, 3, 4, 5 'Cuentas
'            Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo)
        '####
        
        ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
        Case 8: Me.Text2(Index).Text = PonerNombreCCoste(Me.Text1(Index))
    
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
    
    CadB = ObtenerBusqueda(Me, False)

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
''Carga el formulario frmBuscaGrid con los valores correspondientes
'Dim cad As String
'Dim Tabla As String
'Dim Titulo As String
'Dim CargaF As Boolean 'Para saber si se carga el frame o no en el BuscaGrid
'Dim Conexion As Byte
'
'        'Llamamos a al form
'        '##A mano
'        cad = ""
'        If Val(Me.imgCuentas(0).Tag) >= 0 Then
'        'Se llama a Busqueda desde un campo de Cuenta
'            '#A MANO: Porque busca en la tabla Cuentas
'            'de la base de datos de Contabilidad
'            cad = cad & "C�digo|Cuentas|codmacta|T||15�Denominacion|Cuentas|nommacta|T||70�"
'            Tabla = "Cuentas"
'            Titulo = "Cuentas"
'            Conexion = conConta    'Conexi�n a BD: Conta
'            CargaF = True 'Se puede cargar el frame
'        Else
'            'Busqueda de una Fam�lia de Art�culo
'            cad = cad & ParaGrid(Text1(0), 15, "C�digo")
'            cad = cad & ParaGrid(Text1(1), 80, "Denominacion")
'            Tabla = "sfamia"
'            Titulo = "Fam�lia de Art�culos"
'            If vParamAplic.Descriptores Then Titulo = "Categorias Art."
'            Conexion = conAri    'Conexi�n a BD: Aritaxi
'            CargaF = False 'No se carga el frame
'        End If
'
'        If cad <> "" Then
'            Screen.MousePointer = vbHourglass
'            Set frmB = New frmBuscaGrid
'            frmB.vCampos = cad
'            frmB.vTabla = Tabla
'            frmB.vSQL = cadB
'            HaDevueltoDatos = False
'            '###A mano
'            frmB.vDevuelve = "0|1|"
'            frmB.vTitulo = Titulo
'            frmB.vselElem = 1
'            frmB.vConexionGrid = Conexion
'            frmB.vCargaFrame = CargaF
'            '#
'            frmB.Show vbModal
'            Set frmB = Nothing
'            'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'            'tendremos que cerrar el form lanzando el evento
'            If HaDevueltoDatos Then
'                If kCampo < 5 Then PonerFoco Text1(kCampo + 1)
'                If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then
'                    If Not (Val(Me.imgCuentas(0).Tag) >= 0) Then cmdRegresar_Click
'                End If
'            Else   'de ha devuelto datos, es decir NO ha devuelto datos
'                PonerFoco Text1(kCampo)
''                If Modo = 1 Then
''                    MsgBox "No hay ning�n registro en la tabla " & tabla
''                    PonerFoco Text1(0)
''                End If
'            End If
'        End If


    Set frmFam = New frmBasico2
    
    AyudaFamiliasArticulos frmFam, , CadB
    
    Set frmFam = Nothing


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
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim I As Byte
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    'poner la descripcion de las cuentas
    For I = 2 To 7
        Text2(I).Text = PonerNombreCuenta(Text1(I), Modo)
    Next I
        
    ' ---- [19/10/2009] [LAURA]: a�adir campo centro de coste familia
    Me.Text2(8).Text = PonerNombreCCoste(Me.Text1(8))
        
        
    BloquearChecks Me, Modo
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte
Dim I As Integer

    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
    
    '-------------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2) Or Modo = 0
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
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
    
    
    'Poner Boton de Cabecera o Aceptar/Cancelar
    PonerBotonCabecera b Or (Modo = 0)
        
    'Bloquear Registros si modo distinto de Insertar o Modificar
    BloquearText1 Me, Modo
    
    BloquearChecks Me, Modo
        
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n MODO
    PonerOpcionesMenu   'Activar opciones de menu seg�n NIVEL
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
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!Ver, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!Ver, "N")
        
        Toolbar1.Buttons(8).Enabled = Toolbar1.Buttons(8).Enabled And DBLet(Rs!Imprimir, "N")
     End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub






Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
On Error Resume Next

    b = (Modo = 2) Or (Modo = 0) Or (Modo = 1)
    'A�adir
    Toolbar1.Buttons(1).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(3).Enabled = b
    mnEliminar.Enabled = b
    
     '---------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(5).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'VerTodos
    Toolbar1.Buttons(6).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean

    DatosOk = False
    b = CompForm(Me, 1)
    If Not b Then Exit Function
    
    'Comprobar si ya existe el cod de familia en la tabla
    If Modo = 3 Then 'Insertar
        If ExisteCP(Text1(0)) Then b = False
    End If
    
    DatosOk = b
End Function




Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Modo = 3 Or Modo = 4 Then
        Select Case Index
            Case 2, 3, 4, 5, 6, 7 'Cuentas
                Text2(Index).Text = PonerNombreCuenta(Text1(Index), Modo, Text1(1).Text)
                If Text1(Index).Text <> "" And Text2(Index).Text = "" Then Cancel = True
        End Select
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5: mnBuscar_Click
        Case 6: mnVerTodos_Click
        Case 1  'Nuevo
                mnNuevo_Click
        Case 2  'Modificar
                mnModificar_Click
        Case 3  'Borrar
                mnEliminar_Click
        Case 8 'Imprimir listado
            BotonImprimir
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


Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    If b Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codfamia=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim cadParam As String
Dim numParam As Byte
   
    cadFormula = ""
    cadParam = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    cadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = 5
        .Titulo = "Listado Familia de Art�culos"
        .NombreRPT = "rAlmFamArtic.rpt"  'Nombre fichero .rpt a Imprimir
        .Show vbModal
    End With
End Sub


Private Sub TratarCtaContable()
Dim I As Integer
Dim CtaCreadas As String
    For I = 2 To 7
        If Text2(I).Text = vbCrearNuevaCta Then
            If InStr(1, CtaCreadas, Text1(I).Text & "|") = 0 Then
                InsertarCuentaCble Text1(I).Text, "", "", Text1(1).Text
                CtaCreadas = CtaCreadas & Text1(I).Text & "|"
            End If
            Text2(I).Text = Text1(1).Text
        End If
    Next I
End Sub


Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
