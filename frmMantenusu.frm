VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantenusu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantenusu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":7280
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":DAE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantenusu.frx":14344
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameNormal 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.Frame FrameBotonGnral 
         Height          =   705
         Left            =   150
         TabIndex        =   47
         Top             =   0
         Width           =   2655
         Begin VB.CheckBox chkVistaPrevia 
            Caption         =   "Vista previa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3750
            TabIndex        =   48
            Top             =   270
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   240
            TabIndex        =   49
            Top             =   180
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   6
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
                  Object.ToolTipText     =   "Prohibir acceso"
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Object.ToolTipText     =   "Copiar Menus"
                  Object.Tag             =   "0"
               EndProperty
            EndProperty
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         Index           =   1
         ItemData        =   "frmMantenusu.frx":1ABA6
         Left            =   7470
         List            =   "frmMantenusu.frx":1ABB3
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6360
         Width           =   1635
      End
      Begin VB.CommandButton cmdUsu 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5640
         Picture         =   "frmMantenusu.frx":1ABD6
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Prohibir acceso"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdConfigMenu 
         Caption         =   "Configurar menu"
         Height          =   375
         Left            =   7170
         TabIndex        =   38
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   3480
         TabIndex        =   6
         Top             =   900
         Width           =   5655
         Begin VB.ComboBox Combo5 
            Height          =   360
            ItemData        =   "frmMantenusu.frx":21428
            Left            =   960
            List            =   "frmMantenusu.frx":2142A
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1140
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            Height          =   360
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   240
            Width           =   4515
         End
         Begin VB.ComboBox Combo1 
            Height          =   360
            ItemData        =   "frmMantenusu.frx":2142C
            Left            =   960
            List            =   "frmMantenusu.frx":2143C
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   "Skin"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   1170
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   690
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdUsu 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3960
         Picture         =   "frmMantenusu.frx":2146F
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Nuevo usuario"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4440
         Picture         =   "frmMantenusu.frx":27CC1
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Modificar usuario"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4920
         Picture         =   "frmMantenusu.frx":2E513
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar usuario"
         Top             =   5700
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3015
         Left            =   3480
         TabIndex        =   5
         Tag             =   $"frmMantenusu.frx":34D65
         Top             =   3150
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C�digo"
            Object.Width           =   1763
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resumido"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5895
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   10398
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Login"
            Object.Width           =   3352
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   4680
         TabIndex        =   53
         Top             =   2670
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar empresas bloquedas"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarAyuda 
         Height          =   390
         Left            =   8730
         TabIndex        =   56
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ayuda"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Acceso"
         Height          =   255
         Index           =   1
         Left            =   6570
         TabIndex        =   46
         Top             =   6390
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3480
         TabIndex        =   13
         Top             =   690
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Empresas "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame FrameEditorMenus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   120
      TabIndex        =   34
      Top             =   0
      Width           =   9255
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   6015
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10610
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
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
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   36
         Top             =   6360
         Width           =   975
      End
      Begin VB.CommandButton cmdEditorMenus 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   35
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   6360
         Width           =   5055
      End
   End
   Begin VB.Frame FrameUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   60
      TabIndex        =   15
      Top             =   30
      Width           =   9255
      Begin VB.ComboBox Combo6 
         Height          =   360
         ItemData        =   "frmMantenusu.frx":34E06
         Left            =   630
         List            =   "frmMantenusu.frx":34E13
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   5730
         Width           =   2415
      End
      Begin VB.ComboBox Combo4 
         Height          =   360
         ItemData        =   "frmMantenusu.frx":34E38
         Left            =   630
         List            =   "frmMantenusu.frx":34E45
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2820
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   6720
         MaxLength       =   17
         PasswordChar    =   "*"
         TabIndex        =   25
         Text            =   "123456789012345"
         Top             =   4980
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   630
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   4980
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   630
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4260
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   630
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   3540
         Width           =   7695
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2670
         Width           =   1575
      End
      Begin VB.CommandButton cmdFrameUsu 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7080
         TabIndex        =   27
         Top             =   5940
         Width           =   1215
      End
      Begin VB.CommandButton cmdFrameUsu 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5670
         TabIndex        =   26
         Top             =   5940
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6480
         PasswordChar    =   "*"
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2190
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   630
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1410
         Width           =   7725
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frmMantenusu.frx":34E6A
         Left            =   630
         List            =   "frmMantenusu.frx":34E6C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2130
         Width           =   2115
      End
      Begin VB.TextBox Text2 
         Height          =   360
         Index           =   0
         Left            =   630
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   690
         Width           =   1365
      End
      Begin VB.Label Label10 
         Caption         =   "Traer men�s del usuario"
         Height          =   255
         Left            =   630
         TabIndex        =   55
         Top             =   5490
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Skin"
         Height          =   255
         Left            =   630
         TabIndex        =   52
         Top             =   2580
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "mail-password"
         Height          =   255
         Index           =   7
         Left            =   6810
         TabIndex        =   43
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "mail-user"
         Height          =   255
         Index           =   6
         Left            =   630
         TabIndex        =   42
         Top             =   4740
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Servidor SMTP"
         Height          =   255
         Index           =   5
         Left            =   630
         TabIndex        =   41
         Top             =   4020
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "e-mail"
         Height          =   255
         Index           =   4
         Left            =   630
         TabIndex        =   40
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5190
         TabIndex        =   33
         Top             =   240
         Width           =   3345
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   4770
         Top             =   2070
         Width           =   3525
      End
      Begin VB.Label Label4 
         Caption         =   "Confirma Pass."
         Height          =   360
         Index           =   3
         Left            =   5010
         TabIndex        =   32
         Top             =   2670
         Width           =   1605
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   360
         Index           =   2
         Left            =   5040
         TabIndex        =   31
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel"
         Height          =   255
         Left            =   630
         TabIndex        =   30
         Top             =   1890
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre completo"
         Height          =   255
         Index           =   1
         Left            =   630
         TabIndex        =   29
         Top             =   1170
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Login"
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   28
         Top             =   450
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7380
      TabIndex        =   0
      Top             =   5970
      Width           =   975
   End
End
Attribute VB_Name = "frmMantenusu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 105

Dim PrimeraVez As Boolean
Dim Sql As String
Dim I As Integer
Dim UsuarioOrigen As Long


Private Sub cmdConfigMenu_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub


    frmEditorMenus.CodigoActual = CInt(ListView1.SelectedItem.Text)
    frmEditorMenus.Show vbModal

End Sub

Private Sub cmdEditorMenus_Click(Index As Integer)
    If Index = 0 Then
        GuardarMenuUsuario
    End If
    Me.FrameEditorMenus.visible = False
    
    
End Sub


Private Sub cmdFrameUsu_Click(Index As Integer)


    If Index = 0 Then
        If Combo6.ListIndex > 0 Then
            If MsgBox("Va a copiar los menus del usuario " & Trim(Text2(0).Text) & " con los del usuario " & Combo6.Text & vbCrLf & vbCrLf & "� Desea continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                Exit Sub
            End If
        End If
    
        For I = 0 To Text2.Count - 1
            Text2(I).Text = Trim(Text2(I).Text)
            If I < 4 Then
                If Text2(I).Text = "" Then
                    MsgBox Label4(I).Caption & " requerido.", vbExclamation
                    Exit Sub
                End If
            End If
        Next I
        
        If Combo2.ListIndex < 0 Then
            MsgBox "Seleccione un nivel de acceso", vbExclamation
            Exit Sub
        End If
            
        'tipo de skin
        If Combo4.ListIndex < 0 Then
            MsgBox "Seleccione un tipo de skin", vbExclamation
            Exit Sub
        End If
    
        'Password
        If Text2(2).Text <> Text2(3).Text Then
            MsgBox "Password y confirmacion de password no coinciden", vbExclamation
            Exit Sub
        End If
        
        'Ahora vamos con los campos de e-mail
        CadenaDesdeOtroForm = ""
        For I = 4 To 7
            If Text2(I).Text <> "" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "1"
        Next I
        
        If Len(CadenaDesdeOtroForm) > 0 And Len(CadenaDesdeOtroForm) <> 4 Then
            MsgBox "Falta por rellenar correctamente los datos del e-mail.", vbExclamation
            CadenaDesdeOtroForm = ""
            Exit Sub
        End If
        
        'Compruebo que el login es unico
        I = 0
        If UCase(Label6.Caption) = "NUEVO" Then
            Set miRsAux = New ADODB.Recordset
            Sql = "Select login from Usuarios.Usuarios where login='" & Text2(0).Text & "'"
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Sql = ""
            If Not miRsAux.EOF Then Sql = "Ya existe en la tabla usuarios uno con el login: " & miRsAux.Fields(0)
            miRsAux.Close
            Set miRsAux = Nothing
            If Sql <> "" Then
                MsgBox Sql, vbExclamation
                Exit Sub
            End If
            
        Else
            'MODIFICAR
            If FrameUsuario.Tag = 0 Then
                'Estoy modificando un dato normal
                I = CInt(ListView1.SelectedItem.Text)
            Else
                'Estoy agregando un usuario que ya existia en contabi�lidad
                'es decir, le estoy asignando su NIVELUSU de contabilidad
                I = CInt(FrameUsuario.Tag)
            End If
        End If
        
        If Combo6.ListIndex >= 0 Then
            UsuarioOrigen = Combo6.ItemData(Combo6.ListIndex)
        Else
            UsuarioOrigen = 0
        End If
        InsertarModificar I
        
    End If
    
    
    'Cargar usuarios
    If UCase(Label6.Caption) = "NUEVO" Then
        'CargaUsuarios
        CadenaDesdeOtroForm = ""
    Else
        'Pero cargamos el tag como coresponde
        'ListView1.SelectedItem.Tag = Combo2.ItemData(Combo2.ListIndex) & "|" & Text2(1).Text & "|"
        
        If Me.FrameUsuario.Tag <> 0 Then
            CadenaDesdeOtroForm = FrameUsuario.Tag
        Else
            CadenaDesdeOtroForm = ListView1.SelectedItem.Text
        End If
  
    End If
    
    CargaUsuarios
    If CadenaDesdeOtroForm <> "" Then
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Text = CadenaDesdeOtroForm Then
                Set ListView1.SelectedItem = ListView1.ListItems(I)
                Exit For
            End If
        Next I
    End If
    DatosUsusario
    CadenaDesdeOtroForm = ""
    'Para ambos casos
    Me.FrameUsuario.visible = False
    Me.FrameUsuario.Enabled = False
    Me.FrameNormal.visible = True
    Me.FrameNormal.Enabled = True
    
End Sub


Private Sub InsertarModificar(ByVal CodigoUsuario As Integer)
Dim Ant As Integer
Dim Fin As Boolean
Dim SQL2 As String
Dim Excepcion As String

Dim CodUsuarioOrigen As Integer


On Error GoTo EInsertarModificar

    Set miRsAux = New ADODB.Recordset
        
    CodUsuarioOrigen = 0
    If UsuarioOrigen > 0 Then
        CodUsuarioOrigen = DevuelveValor("select codusu from usuarios.usuarios where login = " & DBSet(Combo6.Text, "T"))
    End If
    
    If UCase(Label6.Caption) = "NUEVO" Then
        
        'Nuevo
        Sql = "Select codusu from Usuarios.Usuarios where codusu > 0"
        miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        Ant = 1
        I = 1
        Fin = False
        If miRsAux.EOF Then Fin = True
        While Not Fin
            If miRsAux!codusu - Ant > 0 Then
                'Hay un salto
                I = Ant
                Fin = True
            Else
                Ant = Ant + 1
            End If
            If Not Fin Then
                miRsAux.MoveNext
                If miRsAux.EOF Then
                    Fin = True
                    I = Ant
                End If
            End If
        Wend
        miRsAux.Close

        
        Sql = "INSERT INTO Usuarios.usuarios (codusu, nomusu,  nivelusu, login, passwordpropio,dirfich,skin, solotesoreria) VALUES ("
        Sql = Sql & I
        Sql = Sql & ",'" & Text2(1).Text & "',"
        'Combo
        Sql = Sql & Combo2.ItemData(Combo2.ListIndex) & ",'"
        Sql = Sql & Text2(0).Text & "','"
        Sql = Sql & Text2(3).Text & "',"
        'DIR FICH tiene
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For I = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(I).Text & "|"
            Next I
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        Sql = Sql & CadenaDesdeOtroForm
        
        Sql = Sql & "," & Combo5.ItemData(Combo5.ListIndex) & ","
        Sql = Sql & "0)"
        
        
        
        ' insercion en el menu_usuarios
        SQL2 = "INSERT INTO menus_usuarios (codusu,codigo,aplicacion,ver,creareliminar,modificar,imprimir,especial,expandido) "
        SQL2 = SQL2 & " select " & I & ",codigo, aplicacion, "
        
        ' insertamos sin partir de ning�n usuario
        If UsuarioOrigen <= 0 Then
            Select Case Combo2.ItemData(Combo2.ListIndex)
                Case 0 ' superusuario
                    SQL2 = SQL2 & "1,1,1,1,1,0"
                Case 1 ' administrador
                    SQL2 = SQL2 & "1,1,1,1,1,0"
                Case 2 ' normal
                    SQL2 = SQL2 & "1,1,1,1,1,0"
                Case 3 ' consulta
                    SQL2 = SQL2 & "1,0,0,1,0,0"
            End Select
                    
            SQL2 = SQL2 & " from menus_usuarios "
            SQL2 = SQL2 & " where aplicacion in ('aritaxi','introcon') and codusu = 0"
        ' insertamos partiendo de un usuario
        Else
            SQL2 = SQL2 & " ver, creareliminar, modificar, imprimir, especial, expandido "
            SQL2 = SQL2 & " from menus_usuarios "
            SQL2 = SQL2 & " where aplicacion in ('aritaxi','introcon') and codusu = " & DBSet(CodUsuarioOrigen, "N")
            
        
        End If
        conn.Execute SQL2
        
        
        Excepcion = ""
        ' dependiendo de si es Superusuario, Administrador, Normal o consulta
        Select Case Combo2.ItemData(Combo2.ListIndex)
            Case 0 'superusuario
                
            Case 1 'administrador
                Excepcion = "(1)"
            Case 2 'normal
                Excepcion = "(1,10,12,13,14)"
            Case 3 'consulta
                Excepcion = "(1,10,12,13,14)"
        End Select
        
        If Excepcion <> "" Then
            SQL2 = "update menus_usuarios set ver = 0, creareliminar=0, modificar=0, imprimir = 0, especial= 0, expandido = 0 "
            SQL2 = SQL2 & " where aplicacion in ('aritaxi') and codusu = " & DBSet(I, "N")
            SQL2 = SQL2 & " and (codigo in " & Excepcion
            SQL2 = SQL2 & " or codigo in (select codigo from menus where padre in " & Excepcion & " and aplicacion in ('aritaxi')))"
        
            conn.Execute SQL2
        End If
        
'        If Not vEmpresa.TieneTesoreria Then
'
'            Sql2 = "update menus_usuarios set ver = 0, creareliminar=0, modificar=0, imprimir = 0, especial= 0, expandido = 0 "
'            Sql2 = Sql2 & " where aplicacion in ('ariconta') and codusu = " & DBSet(i, "N")
'            Sql2 = Sql2 & " codigo in (select codigo from menus where aplicacion = 'aritaxi' and tipo = 1)"
'
'            conn.Execute Sql2
'        End If
        
        
    Else
        Sql = "UPDATE Usuarios.Usuarios Set nomusu='" & Text2(1).Text
        
        'Si el combo es administrador compruebo que no fuera en un principio SUPERUSUARIO
        If Combo2.ListIndex = 2 Then
            'Si el combo1 es 3 entonces es super
            If Combo1.ListIndex = 3 Then
                I = 0
            Else
                I = 1
            End If
        Else
            I = Combo2.ItemData(Combo2.ListIndex)
        End If
        Sql = Sql & "' , nivelusu =" & I
        'SQL = SQL & "  , login = '" & Text2(2).Text
        Sql = Sql & "  , passwordpropio = '" & Text2(3).Text & "'"
        
        
        'El e-mail
        If Text2(4).Text = "" Then
            CadenaDesdeOtroForm = "NULL"
        Else
            CadenaDesdeOtroForm = ""
            For I = 4 To 7
                CadenaDesdeOtroForm = CadenaDesdeOtroForm & Text2(I).Text & "|"
            Next I
            CadenaDesdeOtroForm = "'" & CadenaDesdeOtroForm & "'"
        End If
        Sql = Sql & " ,dirfich = " & CadenaDesdeOtroForm
        
        
        Sql = Sql & " ,skin = " & Combo4.ItemData(Combo4.ListIndex)
        
        Sql = Sql & " ,solotesoreria = 0 "
        
        'aqui, en lugar del selecteditem tengo k pasarle el codigo de usuario
        'ya que cuando es nuevo usario y cojo los datos desde otra aplicacion entonces
        'no lo tengo selected y enonces peta
        
        Sql = Sql & " WHERE codusu = " & CodigoUsuario
        
        If UsuarioOrigen <= 0 Then
            SQL2 = "update menus_usuarios set "
            Select Case Combo2.ItemData(Combo2.ListIndex)
                Case 0 'super
                    SQL2 = SQL2 & " ver=1, creareliminar=1, modificar=1, imprimir=1, especial=1"
                Case 1 'administrador
                    SQL2 = SQL2 & " ver=1, creareliminar=1, modificar=1, imprimir=1, especial=1"
                Case 2 'normal
                    SQL2 = SQL2 & " ver=1, creareliminar=1, modificar=1, imprimir=1, especial=1"
                Case 3 'consulta
                    SQL2 = SQL2 & " ver=1, creareliminar=0, modificar=0, imprimir=1, especial=0"
            End Select
            SQL2 = SQL2 & " where codusu = " & DBSet(CodigoUsuario, "N")
            SQL2 = SQL2 & " and aplicacion in ('aritaxi') "
        Else
            'ANTES. MONICA
            'Sql2 = "update menus_usuarios destino, menus_usuarios fuente set "
            'Sql2 = Sql2 & " destino.ver=fuente.ver, destino.creareliminar=fuente.creareliminar, destino.modificar=fuente.modificar, destino.imprimir=fuente.imprimir, destino.especial=fuente.especial"
            'Sql2 = Sql2 & " where destino.codusu = " & DBSet(CodigoUsuario, "N")
            'Sql2 = Sql2 & " and destino.aplicacion in ('ariconta') "
            'Sql2 = Sql2 & " and fuente.codusu = " & DBSet(UsuarioOrigen, "N")
            'Sql2 = Sql2 & " and fuente.aplicacion in ('ariconta') "
            'Sql2 = Sql2 & " and fuente.codigo = destino.codigo "
            SQL2 = "DELETE FROM menus_usuarios WHERE codusu = " & CodigoUsuario
            conn.Execute SQL2
            
            'Preparo el INSERT
            SQL2 = "INSERT INTO menus_usuarios (codusu,codigo,aplicacion,ver,creareliminar,modificar,imprimir,especial,expandido,textovisible,vericono) "
            SQL2 = SQL2 & " SELECT " & CodigoUsuario & ",codigo,aplicacion,ver,creareliminar,modificar,imprimir,especial,expandido,textovisible,vericono FROM menus_usuarios WHERE codusu = " & UsuarioOrigen
            
        End If
        
        conn.Execute SQL2
        
        
        Excepcion = ""
        ' dependiendo de si es Superusuario, Administrador, Normal o consulta
        Select Case Combo2.ItemData(Combo2.ListIndex)
            Case 0 'superusuario
                
            Case 1 'administrador
                Excepcion = "(1)"
            Case 2 'normal
                Excepcion = "(1,10,12,13,14)"
            Case 3 'consulta
                Excepcion = "(1,10,12,13,14)"
        End Select
        
        If Excepcion <> "" Then
            SQL2 = "update menus_usuarios set ver = 0, creareliminar=0, modificar=0, imprimir = 0, especial= 0, expandido = 0"
            SQL2 = SQL2 & " where aplicacion in ('aritaxi') and codusu = " & DBSet(CodigoUsuario, "N")
            SQL2 = SQL2 & " and (codigo in " & Excepcion
            SQL2 = SQL2 & " or codigo in (select codigo from menus where padre in " & Excepcion & " and aplicacion in ('aritaxi')))"
            
            conn.Execute SQL2
        End If
'--
'        If Not vEmpresa.TieneTesoreria Then
'            Sql2 = "update menus_usuarios set ver = 0, creareliminar=0, modificar=0, imprimir = 0, especial= 0, expandido = 0"
'            Sql2 = Sql2 & " where aplicacion in ('ariconta') and codusu = " & DBSet(CodigoUsuario, "N")
'            Sql2 = Sql2 & " and codigo in (select codigo from menus where aplicacion = 'ariconta' and tipo = 1)"
'            conn.Execute Sql2
'        End If
        
        
    End If
    conn.Execute Sql
    
    
    CadenaDesdeOtroForm = ""
    Exit Sub
EInsertarModificar:
    MuestraError Err.Number, "EInsertarModificar"
End Sub



Private Sub cmdUsu_Click(Index As Integer)
Dim K As Integer

    Select Case Index
    Case 0, 1
        limpiar Me
        
        
        
        
        If Index = 0 Then
            'Nuevo usuario
            CargaCombo6 0
            
            Label6.Caption = "NUEVO"
            I = 0 'Para el foco
            
            Combo2.ListIndex = -1
            Combo4.ListIndex = -1
        Else
            
            CargaCombo6 ListView1.SelectedItem
            
            'Modificar0
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            Label6.Caption = "MODIFICAR"
            Set miRsAux = New ADODB.Recordset
            Sql = "Select * from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
            Else
                'LimpiarCamposUsuario
                PonerDatosUsuario
            End If
            I = 1 'Para el foco
            FrameUsuario.Tag = 0  'Marcamos que es una modificacion desde un usuario existente
        End If
        Text2(0).Enabled = (Index = 0)
        
        Me.FrameNormal.visible = False
        Me.FrameNormal.Enabled = False
        Me.FrameUsuario.visible = True
        Me.FrameUsuario.Enabled = True
        Me.FrameEditorMenus.visible = False
        Me.FrameEditorMenus.Enabled = False
        
'        If Not vEmpresa.TieneTesoreria Then Me.ChkSoloTesoreria.Value = 0
        
        Text2(I).SetFocus
        
    Case 2, 3
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        I = vUsu.Codigo Mod 1000
        If I = CInt(ListView1.SelectedItem.Text) Then
            MsgBox "El usuario es el mismo con el que esta trabajando actualmente", vbInformation
            Exit Sub
        End If
        
        If Index = 2 Then
            
            Sql = "El usuario " & ListView1.SelectedItem.SubItems(1) & " ser� eliminado y no tendra acceso a los programas de Ariadna (Ariconta, ariges....) ." & vbCrLf
            Sql = Sql & vbCrLf & "                              �Desea continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Sql = "DELETE from Usuarios.Usuarios where codusu = " & ListView1.SelectedItem.Text
            
        Else
            Sql = "Al usuario " & ListView1.SelectedItem.SubItems(1) & " no le estar� permitido el acceso al programas Aritaxi." & vbCrLf
            Sql = Sql & vbCrLf & "                              �Desea continuar?"
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Sql = "UPDATE Usuarios.usuarios SET nivelaritaxi = -1 WHERE codusu = " & ListView1.SelectedItem.Text
        End If
        Screen.MousePointer = vbHourglass
        
        conn.Execute Sql
    
        '//El codigo siguiente seria mas logico meterlo en el modulo de usuario
        '   pero de momento un saco de cemento
        If Index = 2 Then EliminarAuxiliaresUsuario CInt(ListView1.SelectedItem.Text)
    
        CargaUsuarios
        
        Screen.MousePointer = vbDefault
    
        Me.FrameNormal.visible = True
        Me.FrameNormal.Enabled = True
        Me.FrameUsuario.visible = False
        Me.FrameUsuario.Enabled = False
        Me.FrameEditorMenus.visible = False
        Me.FrameEditorMenus.Enabled = False
    
    End Select

End Sub

Private Sub EliminarAuxiliaresUsuario(codusu As Long)

    On Error GoTo EEliminarAuxiliaresUsuario
    Sql = "DELETE FROM usuarios.usuarioempresasaritaxi where codusu =" & codusu
    conn.Execute Sql
    
    Sql = "DELETE FROM usuarios.appmenususuario where  codusu =" & codusu
    conn.Execute Sql
    
    Sql = "DELETE FROM menus_usuarios where codusu = " & codusu
    conn.Execute Sql
    
    
    Exit Sub
EEliminarAuxiliaresUsuario:
    MuestraError Err.Number, "Eliminar Auxiliares Usuario"

End Sub

Private Sub PonerDatosUsuario()
        
     Text2(0).Text = miRsAux!Login
     Text2(1).Text = miRsAux!nomusu
     Text2(2).Text = miRsAux!passwordpropio
     Text2(3).Text = miRsAux!passwordpropio
     I = miRsAux!nivelusu

    Select Case I
        Case 0
            I = 3
        Case 1
            I = 2
        Case 2
            I = 1
        Case 3
            I = 0
    End Select


    Combo2.ListIndex = I
    
    PosicionarCombo Combo4, miRsAux!Skin
     
     'Cargamos los datos del correo e-mail
     Sql = Trim(DBLet(miRsAux!Dirfich, "T"))
     If Sql <> "" Then
         For I = 1 To 4
             Text2(3 + I).Text = RecuperaValor(Sql, I)
         Next I
     End If
     
'     Me.ChkSoloTesoreria.Value = DBLet(miRsAux!SoloTesoreria)

End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo3_Click(Index As Integer)
    If Not PrimeraVez Then DatosUsusario
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaUsuarios
    End If
    FrameEditorMenus.visible = False
    LeerEditorMenus
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    Me.Icon = frmppal.Icon

    PrimeraVez = True
    
    ' Botonera Principal
    With Me.Toolbar1
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        .Buttons(1).Image = 3
        .Buttons(2).Image = 4
        .Buttons(3).Image = 5
        .Buttons(5).Image = 14
        .Buttons(6).Image = 11
    End With
    
    With Me.Toolbar2
        .HotImageList = frmppal.imgListComun_OM
        .DisabledImageList = frmppal.imgListComun_BN
        .ImageList = frmppal.imgListComun1
        .Buttons(1).Image = 28
    End With
    
'    ' La Ayuda
'    With Me.ToolbarAyuda
'        .ImageList = frmppal.ImgListComun1
'        .Buttons(1).Image = 26
'    End With
    
    CargaCombo
    
    
'    Me.ChkSoloTesoreria.visible = vEmpresa.TieneTesoreria
'    Me.ChkSoloTesoreria.Enabled = vEmpresa.TieneTesoreria
    
    Me.ListView1.SmallIcons = ImageList1
    Me.ListView2.SmallIcons = ImageList1
    Me.FrameUsuario.visible = False
    Me.FrameNormal.Enabled = True
'    LeerDatosCombo True

    PonerModoUsuarioGnral 0, "aritaxi"

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
        
        Toolbar1.Buttons(5).Enabled = Toolbar1.Buttons(5).Enabled And DBLet(Rs!especial, "N")
        Toolbar1.Buttons(6).Enabled = Toolbar1.Buttons(6).Enabled And DBLet(Rs!especial, "N")
        
    End If
    
    Rs.Close
    Set Rs = Nothing
    
End Sub


'Private Sub LeerDatosCombo(Leer As Boolean)
'Dim Cad2 As String
'
'    On Error GoTo ELe
'    If Leer Then
'
'        Combo3(1).ListIndex = 0
'        I = vControl.UltAccesoBDs  'RecuperaValor(CadenaControl, 6)
'        Combo3(1).ListIndex = I
'    Else
'        'GUARDAR
'        vControl.UltAccesoBDs = Combo3(1).ListIndex
'        vControl.Grabar
'
'            CadenaControl = InsertaValor(CadenaControl, 6, Combo3(1).ListIndex)
'
'    End If
'    Exit Sub
'ELe:
'    Err.Clear
'End Sub

Private Sub CargaUsuarios()
Dim Itm As ListItem

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    '                               Aquellos usuarios k tengan nivel usu -1 NO son de conta
    '  QUitamos codusu=0 pq es el usuario ROOT
    Sql = "Select * from Usuarios.Usuarios where nivelaritaxi >=0 "
    
    ' solo vemos root si somos root
    If vUsu.Login = "root" Then
        Sql = Sql & " and codusu >= 0 order by codusu"
    Else
        Sql = Sql & " and codusu > 0 order by codusu"
    End If
    
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set Itm = ListView1.ListItems.Add
        Itm.Text = miRsAux!codusu
        Itm.SubItems(1) = miRsAux!Login
        If miRsAux!nivelusu = 0 Then
            Itm.SmallIcon = 4
        Else
            Itm.SmallIcon = 5
        End If
        'Nombre y nivel de usuario
        Sql = miRsAux!nivelusu & "|" & miRsAux!nomusu & "|"
        Itm.Tag = Sql
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
        DatosUsusario
    End If

End Sub



Private Sub DatosUsusario()
Dim ItmX As ListItem
On Error GoTo EDatosUsu

    ListView2.ListItems.Clear
    If ListView1.SelectedItem Is Nothing Then
        Text4.Text = ""
        Combo1.ListIndex = -1
        Combo5.ListIndex = -1
        Exit Sub
    End If
    
    
    Text4.Text = RecuperaValor(ListView1.SelectedItem.Tag, 2)
    'NIVEL
    Sql = RecuperaValor(ListView1.SelectedItem.Tag, 1)
    '                           COMBO                      en Bd
    '                       0.- Consulta                     3
    '                       1.- Normal                       2
    '                       2.- Administrador                1
    '                       3.- SuperUsuario (root)          0
    If Not IsNumeric(Sql) Then Sql = 3
    Select Case Val(Sql)
    Case 2
        Combo1.ListIndex = 1
    Case 1
        Combo1.ListIndex = 2
    Case 0
        Combo1.ListIndex = 3
    Case Else
        Combo1.ListIndex = 0
    End Select
    
    
    Sql = DevuelveValor("select skin from usuarios.usuarios where codusu = " & ListView1.SelectedItem.Text)
    PosicionarCombo Combo5, CInt(Val(Sql))
    
    
    Sql = "select empresasaritaxi.codempre,nomempre,nomresum,usuarioempresasaritaxi.codempre bloqueada from usuarios.empresasaritaxi left join usuarios.usuarioempresasaritaxi on "
    Sql = Sql & " empresasaritaxi.codempre = usuarioempresasaritaxi.codempre And (usuarioempresasaritaxi.codusu = " & ListView1.SelectedItem.Text & " Or codusu Is Null)"
    
    
    If Combo3(1).ListIndex > 0 Then
            Sql = Sql & " WHERE "
        If Combo3(1).ListIndex = 1 Then Sql = Sql & " NOT "
        Sql = Sql & " (usuarioempresasaritaxi.codempre is null) "
    End If
    
    '[Monica] s�lo empresas de aritaxi nuevas
    If InStr(1, Sql, "WHERE") <> 0 Then
        Sql = Sql & " and empresasaritaxi.aritaxi like 'aritaxi%' "
    Else
        Sql = Sql & " WHERE empresasaritaxi.aritaxi like 'aritaxi%' "
    End If
    
    'EL orden
    Sql = Sql & " ORDER BY empresasaritaxi.codempre"
    
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRsAux.EOF
        Set ItmX = ListView2.ListItems.Add
        ItmX.Text = miRsAux.Fields(0)
        ItmX.SubItems(1) = miRsAux!nomempre
        ItmX.SubItems(2) = miRsAux!nomresum
        If miRsAux.Fields(0) > 100 Then
            ItmX.ForeColor = &H808080
            ItmX.ListSubItems(1).ForeColor = &H808080
            ItmX.ListSubItems(2).ForeColor = &H808080
        End If
        
        If IsNull(miRsAux!bloqueada) Then
            ItmX.SmallIcon = 1
        Else
            ItmX.SmallIcon = 2
        End If
    
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    
    Exit Sub
EDatosUsu:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    LeerDatosCombo False
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus(Index As Integer)
Dim AsignarDatos As Boolean

    Text2(Index).Text = Trim(Text2(Index).Text)
    If Text2(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        If UCase(Label6.Caption) = "NUEVO" Then
        
            'Si es nuevo entonces, primero compruebo que no existe el login
            'Si existe, y el usuario tiene nivel conta >=0 entonces
            ' existe en la conta. Si existe pero el nivel conta es -1 entonces
            'lo que hacemos es ponerle los datos y que cambie la opcion de nivel usu
            Sql = "Select * from usuarios.usuarios where login='" & Text2(0).Text & "'"
            Set miRsAux = New ADODB.Recordset
            miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRsAux.EOF Then
                'Tiene nivel usu
                If miRsAux!nivelusu > 0 Then
                    MsgBox "El usuario ya existe para la contabilidad", vbExclamation
                    LimpiarCamposUsuario
                    Text2(0).SetFocus
                    
                Else
                    If miRsAux!codusu = 0 Then
                        MsgBox "Esta intentando modificar datos del usuario ADMINISTRADOR", vbCritical
                        AsignarDatos = False
                    Else
                        Sql = "El usuario existe para otras aplicaciones de Ariadna Software." & vbCrLf
                        Sql = Sql & "�Desea agregarlo como usuario a la contabilidad?"
                        If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then AsignarDatos = True
                    End If
                    If AsignarDatos Then
                        PonerDatosUsuario
                        'Ahora pongo el label y el campo a disbled
                        Text2(1).SetFocus
                        Label6.Caption = "MODIFICAR"
                        Text2(0).Enabled = False
                        FrameUsuario.Tag = miRsAux!codusu 'Pongo el frame al codigo ndel usuario
                    Else
                        LimpiarCamposUsuario
                        Text2(0).SetFocus
                    End If
                End If
            End If
            miRsAux.Close
        End If
    End If
    
End Sub

Private Sub LimpiarCamposUsuario()
    For I = 0 To 7
        Text2(I).Text = ""
    Next I
End Sub
Private Sub LeerEditorMenus()
    On Error GoTo ELeerEditorMenus
    cmdConfigMenu.visible = vUsu.Nivel < 1
      
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub







Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' insertar
'            UsuarioOrigen = 0
            cmdUsu_Click (0)
        Case 2 ' modificar
            cmdUsu_Click (1)
        Case 3 ' eliminar
            cmdUsu_Click (2)
        Case 5 ' prohibir acceso
            cmdUsu_Click (3)
        Case 6 ' copiar menus
            If Not ListView1.SelectedItem Is Nothing Then
'                UsuarioOrigen = ListView1.SelectedItem
                cmdUsu_Click (0)
            End If
        Case Else
        
    End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un usuario", vbExclamation
        Exit Sub
    End If

    frmMensajes.OpcionMensaje = 30
    frmMensajes.Parametros = ListView1.SelectedItem.Text
    frmMensajes.Show vbModal
    
    DatosUsusario

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
If Node.Children > 0 Then Recursivo2 Node.Child, Node.Checked
End Sub


Private Sub CheckarNodo(N As Node, Valor As Boolean)
Dim no As Node
    Set no = N.LastSibling
    Do
        N.Checked = Valor
        If N.Children > 0 Then CheckarNodo N, Valor
        If N.Next <> no.LastSibling Then Set N = N.Next
    Loop Until no = N
End Sub

Private Sub Recursivo2(ByVal Nod As Node, Valor As Boolean)
Dim nx As Node
Dim Aux

    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
        'aux = nx.Root
        'aux = nx.Parent
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then Recursivo2 nx.Child, Valor
        nx.Checked = Valor
      End If
    Set nx = Nothing
End Sub


Private Sub GuardarMenuUsuario()
    Sql = "DELETE from usuarios.appmenusUsuario where aplicacion='Conta' AND codusu =" & ListView1.SelectedItem.Text
    conn.Execute Sql
    
    I = 0
    Sql = "INSERT INTO usuarios.appmenususuario (aplicacion, codusu, codigo, tag) VALUES ('Conta'," & ListView1.SelectedItem.Text & ","
    RecursivoBD TreeView1.Nodes(1)
End Sub

Private Sub InsertaBD(vtag As String)
Dim C As String
    I = I + 1
    'SQL = "INSERT INTO appmenususuario (aplicacion, codusu, codigo, tag)
    C = Sql & I & ",'" & vtag & "')"
    conn.Execute C
End Sub


Private Sub RecursivoBD(ByVal Nod As Node)
Dim nx As Node
Dim Aux

    
    
    Set nx = Nod.FirstSibling
    While nx <> Nod.LastSibling
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
        Set nx = nx.Next
    Wend
    
    If nx = Nod.LastSibling Then
        If nx.Children > 0 Then
            If nx.Checked Then RecursivoBD nx.Child
        End If
        If Not nx.Checked Then InsertaBD nx.Tag
      End If
    Set nx = Nothing
End Sub

Private Sub CargaCombo()
    
    'nivel
    Combo2.Clear
    
    Combo2.AddItem "Consulta"
    Combo2.ItemData(Combo2.NewIndex) = 3
    
    Combo2.AddItem "Normal"
    Combo2.ItemData(Combo2.NewIndex) = 2
    
    Combo2.AddItem "Administrador"
    Combo2.ItemData(Combo2.NewIndex) = 1
    
    Combo2.AddItem "Superusuario"
    Combo2.ItemData(Combo2.NewIndex) = 0


    '3 ID_OPTIONS_STYLEBLACK2010
    '2 S_STYLESILVER2010
    '1ID_OPTIONS_STYLEBLUE2010

    'skin
    Combo5.Clear
    
    Combo5.AddItem "Office 2010 Blue"
    Combo5.ItemData(Combo5.NewIndex) = 1
    
    Combo5.AddItem "Office 2010 Silver"
    Combo5.ItemData(Combo5.NewIndex) = 2
    
    Combo5.AddItem "Office 2010 Black"
    Combo5.ItemData(Combo5.NewIndex) = 3
    
    
    'skin
    Combo4.Clear
    
    Combo4.AddItem "Office 2010 Blue"
    Combo4.ItemData(Combo4.NewIndex) = 1
    
    Combo4.AddItem "Office 2010 Silver"
    Combo4.ItemData(Combo4.NewIndex) = 2
    
    Combo4.AddItem "Office 2010 Black"
    Combo4.ItemData(Combo4.NewIndex) = 3

    
End Sub


Private Sub CargaCombo6(Usuario As Integer)
Dim Sql As String
Dim Rs As ADODB.Recordset

    'skin
    Combo6.Clear
    
    Sql = "select codusu, login from usuarios.usuarios where codusu <> " & DBSet(Usuario, "N") & " and login <> 'root' and nivelusu > -1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Combo6.AddItem "Ninguno"
    Combo6.ItemData(Combo6.NewIndex) = 0
    
    While Not Rs.EOF
        Combo6.AddItem Rs.Fields(1).Value
        Combo6.ItemData(Combo6.NewIndex) = Rs.Fields(0).Value
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
End Sub



