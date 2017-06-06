VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoOfer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11310
   Icon            =   "frmListadoOfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameClienInactivos 
      Height          =   7005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11145
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   6600
         TabIndex        =   146
         Top             =   1680
         Width           =   4215
         Begin VB.Frame Frame5 
            Caption         =   "e-Mail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   600
            TabIndex        =   15
            Top             =   1680
            Width           =   2535
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Comercial"
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
               TabIndex        =   150
               Top             =   460
               Width           =   2265
            End
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
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
               TabIndex        =   149
               Top             =   210
               Value           =   -1  'True
               Width           =   2265
            End
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   64
            Left            =   180
            MaxLength       =   6
            TabIndex        =   13
            Top             =   860
            Width           =   615
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   64
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   147
            Text            =   "Text5"
            Top             =   860
            Width           =   3375
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
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
            Left            =   180
            TabIndex        =   14
            Top             =   1395
            Width           =   2415
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   10
            Left            =   180
            TabIndex        =   148
            Top             =   585
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   40
            Left            =   840
            Top             =   580
            Width           =   240
         End
      End
      Begin VB.Frame FrameImpClien 
         Caption         =   "Imprimir clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1110
         Left            =   600
         TabIndex        =   12
         Top             =   5850
         Visible         =   0   'False
         Width           =   2805
         Begin VB.OptionButton OptCliTodos 
            Caption         =   "Todos"
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
            TabIndex        =   124
            Top             =   795
            Width           =   1605
         End
         Begin VB.OptionButton OptCliSinMante 
            Caption         =   "Sin mantenimiento"
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
            TabIndex        =   123
            Top             =   510
            Width           =   2325
         End
         Begin VB.OptionButton OptCliConMante 
            Caption         =   "Con mantenimiento"
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
            TabIndex        =   122
            Top             =   240
            Width           =   2325
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2940
         Left            =   480
         TabIndex        =   110
         Top             =   2925
         Visible         =   0   'False
         Width           =   5925
         Begin VB.TextBox txtNombre 
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
            Index           =   57
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   125
            Text            =   "Text5"
            Top             =   1995
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   57
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   10
            Top             =   1995
            Width           =   855
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2595
            Width           =   4095
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   56
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   9
            Top             =   1500
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   56
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   118
            Text            =   "Text5"
            Top             =   1500
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   55
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   8
            Top             =   1130
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   55
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   117
            Text            =   "Text5"
            Top             =   1130
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   54
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   7
            Top             =   615
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   54
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   112
            Text            =   "Text5"
            Top             =   615
            Width           =   3615
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   53
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   53
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   111
            Text            =   "Text5"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   34
            Left            =   960
            Top             =   1995
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Situación"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   73
            Left            =   120
            TabIndex        =   126
            Top             =   1785
            Width           =   975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "A la atención de:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   71
            Left            =   120
            TabIndex        =   121
            Top             =   2355
            Width           =   1785
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   70
            Left            =   300
            TabIndex        =   120
            Top             =   1470
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   33
            Left            =   960
            Top             =   1500
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   69
            Left            =   300
            TabIndex        =   119
            Top             =   1125
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   32
            Left            =   960
            Top             =   1130
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "CPostal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   68
            Left            =   120
            TabIndex        =   116
            Top             =   885
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   67
            Left            =   300
            TabIndex        =   115
            Top             =   585
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   31
            Left            =   960
            Top             =   615
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   66
            Left            =   300
            TabIndex        =   114
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Actividad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   57
            Left            =   120
            TabIndex        =   113
            Top             =   0
            Width           =   990
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   30
            Left            =   960
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   4230
         MaxLength       =   10
         TabIndex        =   18
         Top             =   3360
         Width           =   1230
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3360
         Width           =   1230
      End
      Begin VB.CommandButton cmdAceptarClienInac 
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
         Left            =   3870
         TabIndex        =   16
         Top             =   6240
         Width           =   1135
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   5
         Left            =   5100
         TabIndex        =   17
         Top             =   6240
         Width           =   1135
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   27
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1260
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   27
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   28
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   1635
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1635
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   29
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2200
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   3
         Top             =   2200
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   30
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2580
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   4
         TabIndex        =   4
         Top             =   2580
         Width           =   855
      End
      Begin VB.Frame frameCliexFacturas 
         Caption         =   "Desde / hasta facturas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   6480
         TabIndex        =   290
         Top             =   1080
         Width           =   4455
         Begin VB.ComboBox cboTipomov 
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
            ItemData        =   "frmListadoOfer.frx":000C
            Left            =   920
            List            =   "frmListadoOfer.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   291
            Top             =   840
            Width           =   1875
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   104
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   294
            Top             =   3240
            Width           =   1230
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   102
            Left            =   960
            MaxLength       =   7
            TabIndex        =   292
            Text            =   "wwwwwww"
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   103
            Left            =   2640
            MaxLength       =   7
            TabIndex        =   293
            Top             =   2160
            Width           =   1365
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   105
            Left            =   2640
            MaxLength       =   10
            TabIndex        =   295
            Top             =   3240
            Width           =   1230
         End
         Begin VB.Image imgClearCmbTipomov 
            Height          =   240
            Left            =   2880
            Picture         =   "frmListadoOfer.frx":0010
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movimiento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   12
            Left            =   120
            TabIndex        =   302
            Top             =   600
            Width           =   1785
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   31
            Left            =   1875
            Top             =   3000
            Width           =   240
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Fact."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   301
            Top             =   2850
            Width           =   1215
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nº Factura"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   11
            Left            =   240
            TabIndex        =   300
            Top             =   1680
            Width           =   1140
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   10
            Left            =   960
            TabIndex        =   299
            Top             =   1920
            Width           =   600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   9
            Left            =   2640
            TabIndex        =   298
            Top             =   1920
            Width           =   570
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   3
            Left            =   1080
            TabIndex        =   297
            Top             =   3030
            Width           =   600
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   2
            Left            =   2640
            TabIndex        =   296
            Top             =   3030
            Width           =   570
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   32
            Left            =   3420
            Top             =   3000
            Width           =   240
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   44
         Left            =   3255
         TabIndex        =   32
         Top             =   3360
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   12
         Left            =   3930
         Top             =   3375
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   43
         Left            =   780
         TabIndex        =   31
         Top             =   3360
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   1440
         Top             =   3380
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inactividad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   36
         Left            =   600
         TabIndex        =   30
         Top             =   3120
         Width           =   1905
      End
      Begin VB.Label Label8 
         Caption         =   "Clientes Inactivos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   600
         TabIndex        =   29
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   9
         Left            =   1440
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   42
         Left            =   600
         TabIndex        =   28
         Top             =   1035
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   41
         Left            =   780
         TabIndex        =   27
         Top             =   1260
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   10
         Left            =   1440
         Top             =   1635
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   40
         Left            =   780
         TabIndex        =   26
         Top             =   1605
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   11
         Left            =   1440
         Top             =   2200
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   39
         Left            =   600
         TabIndex        =   25
         Top             =   1935
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   38
         Left            =   780
         TabIndex        =   24
         Top             =   2205
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   12
         Left            =   1440
         Top             =   2580
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   37
         Left            =   780
         TabIndex        =   23
         Top             =   2550
         Width           =   570
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEnvioMail 
      Height          =   1215
      Left            =   0
      TabIndex        =   303
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   360
         TabIndex        =   304
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Preparando datos envio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   360
         TabIndex        =   305
         Top             =   840
         Width           =   5805
      End
   End
   Begin VB.Frame FrameEstCliente 
      Height          =   4395
      Left            =   0
      TabIndex        =   329
      Top             =   0
      Width           =   7035
      Begin VB.ComboBox Combo2 
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
         Left            =   1620
         TabIndex        =   334
         Text            =   "Combo1"
         Top             =   2880
         Width           =   2925
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   0
         Left            =   5550
         TabIndex        =   336
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarEstCliente 
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
         Left            =   4320
         TabIndex        =   335
         Top             =   3690
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   331
         Top             =   1470
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   346
         Text            =   "Text5"
         Top             =   1485
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   330
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   345
         Text            =   "Text5"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   333
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   332
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   450
         TabIndex        =   344
         Top             =   2610
         Width           =   1785
      End
      Begin VB.Label Label9 
         Caption         =   "Detalle Facturación Clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   44
         Left            =   480
         TabIndex        =   343
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   43
         Left            =   630
         TabIndex        =   342
         Top             =   1425
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   5
         Left            =   1320
         Top             =   1485
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   42
         Left            =   630
         TabIndex        =   341
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   41
         Left            =   480
         TabIndex        =   340
         Top             =   795
         Width           =   765
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   4
         Left            =   1320
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   4
         Left            =   3030
         TabIndex        =   339
         Top             =   2130
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   3735
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   338
         Top             =   1890
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   2
         Left            =   630
         TabIndex        =   337
         Top             =   2130
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1320
         Top             =   2160
         Width           =   240
      End
   End
   Begin VB.Frame FrameEstVentasFam 
      Height          =   5925
      Left            =   480
      TabIndex        =   261
      Top             =   0
      Width           =   7035
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         Height          =   945
         Left            =   270
         TabIndex        =   326
         Top             =   2400
         Visible         =   0   'False
         Width           =   4455
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
            Left            =   1350
            TabIndex        =   328
            Text            =   "Combo1"
            Top             =   540
            Width           =   2925
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Movimiento"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   0
            Left            =   210
            TabIndex        =   327
            Top             =   270
            Width           =   1785
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Clasificado por "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   480
         TabIndex        =   323
         Top             =   4960
         Width           =   2565
         Begin VB.OptionButton OptPorCliente 
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
            Height          =   240
            Left            =   1320
            TabIndex        =   325
            Top             =   280
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton OptPorFamilia 
            Caption         =   "Familia"
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
            TabIndex        =   324
            Top             =   280
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   99
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   267
         Top             =   2040
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   98
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   266
         Top             =   2040
         Width           =   1350
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   96
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   281
         Text            =   "Text5"
         Top             =   1020
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   96
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   264
         Top             =   1020
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   97
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   280
         Text            =   "Text5"
         Top             =   1395
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   97
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   265
         Top             =   1395
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarEstVentas 
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
         Left            =   4320
         TabIndex        =   274
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   17
         Left            =   5520
         TabIndex        =   275
         Top             =   5160
         Width           =   975
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   240
         TabIndex        =   262
         Top             =   2430
         Width           =   6495
         Begin VB.CheckBox chkDatosAlbaranes 
            Caption         =   "Datos albaranes"
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
            Left            =   3690
            TabIndex        =   271
            Top             =   1200
            Width           =   2325
         End
         Begin VB.CheckBox chkDetallaArticulo 
            Caption         =   "Detalla articulo"
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
            Left            =   240
            TabIndex        =   270
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Frame FrameDetallaArticulo 
            BorderStyle     =   0  'None
            Caption         =   "Frame10"
            Height          =   975
            Left            =   240
            TabIndex        =   306
            Top             =   1500
            Visible         =   0   'False
            Width           =   6135
            Begin VB.TextBox txtNombre 
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
               Index           =   113
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   310
               Text            =   "Text5"
               Top             =   630
               Width           =   3735
            End
            Begin VB.TextBox txtCodigo 
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
               Index           =   113
               Left            =   1140
               MaxLength       =   16
               TabIndex        =   273
               Top             =   630
               Width           =   1095
            End
            Begin VB.TextBox txtNombre 
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
               Index           =   112
               Left            =   2280
               Locked          =   -1  'True
               TabIndex        =   307
               Text            =   "Text5"
               Top             =   240
               Width           =   3735
            End
            Begin VB.TextBox txtCodigo 
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
               Index           =   112
               Left            =   1140
               MaxLength       =   16
               TabIndex        =   272
               Top             =   240
               Width           =   1095
            End
            Begin VB.Image imgBuscarOfer 
               Height          =   240
               Index           =   59
               Left            =   840
               Top             =   630
               Width           =   240
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
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
               Index           =   37
               Left            =   150
               TabIndex        =   311
               Top             =   630
               Width           =   570
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Artículo"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Index           =   36
               Left            =   0
               TabIndex        =   309
               Top             =   0
               Width           =   810
            End
            Begin VB.Image imgBuscarOfer 
               Height          =   240
               Index           =   58
               Left            =   840
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
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
               Index           =   35
               Left            =   150
               TabIndex        =   308
               Top             =   240
               Width           =   600
            End
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   101
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   269
            Top             =   735
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   101
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   276
            Text            =   "Text5"
            Top             =   735
            Width           =   4125
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   100
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   268
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   100
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   263
            Text            =   "Text5"
            Top             =   360
            Width           =   4125
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   27
            Left            =   390
            TabIndex        =   279
            Top             =   735
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   55
            Left            =   1080
            Top             =   735
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   26
            Left            =   390
            TabIndex        =   278
            Top             =   360
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   54
            Left            =   1080
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   25
            Left            =   240
            TabIndex        =   277
            Top             =   90
            Width           =   780
         End
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   30
         Left            =   3720
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   91
         Left            =   630
         TabIndex        =   288
         Top             =   2040
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   90
         Left            =   480
         TabIndex        =   287
         Top             =   1800
         Width           =   630
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   29
         Left            =   1335
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   89
         Left            =   3030
         TabIndex        =   286
         Top             =   2040
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   52
         Left            =   1320
         Top             =   1020
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   30
         Left            =   480
         TabIndex        =   284
         Top             =   795
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   29
         Left            =   630
         TabIndex        =   283
         Top             =   1020
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   53
         Left            =   1320
         Top             =   1395
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   28
         Left            =   630
         TabIndex        =   282
         Top             =   1395
         Width           =   570
      End
      Begin VB.Label Label9 
         Caption         =   "Ventas por Familia / Artículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   31
         Left            =   600
         TabIndex        =   285
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame FrameCierreCaja 
      Height          =   3735
      Left            =   0
      TabIndex        =   221
      Top             =   0
      Width           =   6315
      Begin VB.Frame FrameAgrupar 
         Caption         =   "Agrupar por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1000
         Left            =   600
         TabIndex        =   232
         Top             =   2160
         Width           =   2415
         Begin VB.OptionButton optForpago 
            Caption         =   "Tipo de pago"
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
            TabIndex        =   225
            Top             =   620
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton optForpago 
            Caption         =   "Forma de pago"
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
            TabIndex        =   224
            Top             =   320
            Width           =   2085
         End
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   88
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   222
         Top             =   1560
         Width           =   1350
      End
      Begin VB.CommandButton cmdAceptarCierre 
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
         Left            =   3240
         TabIndex        =   226
         Top             =   2785
         Width           =   1135
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   15
         Left            =   4440
         TabIndex        =   227
         Top             =   2785
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   89
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   223
         Top             =   1560
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   3
         Left            =   3300
         TabIndex        =   231
         Top             =   1560
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   25
         Left            =   1480
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label10 
         Caption         =   "Cierre de Caja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   230
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   229
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   2
         Left            =   780
         TabIndex        =   228
         Top             =   1560
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   26
         Left            =   3960
         Top             =   1560
         Width           =   240
      End
   End
   Begin VB.Frame FrameFacReimprimir 
      Height          =   4455
      Left            =   240
      TabIndex        =   203
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox chkFormatoTPV 
         Caption         =   "Formato factura TPV"
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
         Left            =   480
         TabIndex        =   289
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox chk_duplicado 
         Caption         =   "Duplicado"
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
         Left            =   480
         TabIndex        =   219
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   86
         Left            =   4365
         MaxLength       =   10
         TabIndex        =   208
         Top             =   2880
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   84
         Left            =   4380
         MaxLength       =   7
         TabIndex        =   206
         Top             =   2172
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   83
         Left            =   1695
         MaxLength       =   7
         TabIndex        =   205
         Text            =   "wwwwwww"
         Top             =   2172
         Width           =   1365
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   85
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   207
         Top             =   2880
         Width           =   1350
      End
      Begin VB.CommandButton cmdAceptarReimpFac 
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
         Left            =   3750
         TabIndex        =   209
         Top             =   3840
         Width           =   1135
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   14
         Left            =   4950
         TabIndex        =   210
         Top             =   3840
         Width           =   1135
      End
      Begin VB.ComboBox cboTipomov 
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
         ItemData        =   "frmListadoOfer.frx":059A
         Left            =   1695
         List            =   "frmListadoOfer.frx":059C
         Style           =   2  'Dropdown List
         TabIndex        =   204
         Top             =   1425
         Width           =   2475
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   24
         Left            =   4080
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   8
         Left            =   3450
         TabIndex        =   218
         Top             =   2895
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   7
         Left            =   690
         TabIndex        =   217
         Top             =   2895
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   6
         Left            =   3450
         TabIndex        =   216
         Top             =   2145
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   5
         Left            =   690
         TabIndex        =   215
         Top             =   2145
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   480
         TabIndex        =   214
         Top             =   1860
         Width           =   1140
      End
      Begin VB.Label Label14 
         Caption         =   "Reimprimir Facturas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   480
         TabIndex        =   213
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fact."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   480
         TabIndex        =   212
         Top             =   2595
         Width           =   1215
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   23
         Left            =   1400
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   211
         Top             =   1140
         Width           =   1785
      End
   End
   Begin VB.Frame FrameFacRectif 
      Height          =   4455
      Left            =   720
      TabIndex        =   178
      Top             =   480
      Width           =   5715
      Begin VB.TextBox txtCodigo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   87
         Left            =   480
         MaxLength       =   72
         MultiLine       =   -1  'True
         TabIndex        =   186
         Top             =   2760
         Width           =   4875
      End
      Begin VB.ComboBox cboTipomov 
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
         ItemData        =   "frmListadoOfer.frx":059E
         Left            =   2160
         List            =   "frmListadoOfer.frx":05A0
         Style           =   2  'Dropdown List
         TabIndex        =   183
         Top             =   1185
         Width           =   1875
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   11
         Left            =   4200
         TabIndex        =   188
         Top             =   3720
         Width           =   1135
      End
      Begin VB.CommandButton cmdAceptarFacRect 
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
         Left            =   3000
         TabIndex        =   187
         Top             =   3720
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   72
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   185
         Top             =   2115
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   71
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   184
         Top             =   1725
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Motivo"
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
         Index           =   82
         Left            =   480
         TabIndex        =   220
         Top             =   2520
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movimiento"
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
         Index           =   79
         Left            =   480
         TabIndex        =   182
         Top             =   1200
         Width           =   1605
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   17
         Left            =   1860
         Top             =   2115
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   77
         Left            =   480
         TabIndex        =   181
         Top             =   2115
         Width           =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Factura a Rectificar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   5
         Left            =   480
         TabIndex        =   180
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   76
         Left            =   480
         TabIndex        =   179
         Top             =   1725
         Width           =   1080
      End
   End
   Begin VB.Frame FrameGenAlbCom 
      Height          =   4455
      Left            =   240
      TabIndex        =   76
      Top             =   240
      Width           =   6315
      Begin VB.TextBox txtCodigo 
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
         Index           =   48
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   79
         Top             =   2520
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   49
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   80
         Top             =   3000
         Width           =   1350
      End
      Begin VB.CommandButton cmdAceptarAlbCom 
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
         Left            =   3720
         TabIndex        =   81
         Top             =   3840
         Width           =   1135
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   7
         Left            =   4980
         TabIndex        =   82
         Top             =   3840
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   47
         Left            =   900
         MaxLength       =   4
         TabIndex        =   78
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   47
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Albarán"
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
         Index           =   61
         Left            =   270
         TabIndex        =   96
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   3
         Left            =   270
         TabIndex        =   95
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Alb."
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
         Index           =   60
         Left            =   270
         TabIndex        =   94
         Top             =   3000
         Width           =   1050
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1440
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el Albaran de compra: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   59
         Left            =   270
         TabIndex        =   84
         Top             =   1200
         Width           =   5970
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador del Albaran"
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
         Index           =   58
         Left            =   270
         TabIndex        =   83
         Top             =   1650
         Width           =   2070
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   27
         Left            =   600
         Top             =   1920
         Width           =   240
      End
   End
   Begin VB.Frame FramePteRecibir 
      Height          =   5205
      Left            =   480
      TabIndex        =   159
      Top             =   240
      Width           =   7035
      Begin VB.Frame Frame7 
         Caption         =   "Ordenar por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   940
         Left            =   600
         TabIndex        =   175
         Top             =   3960
         Width           =   2055
         Begin VB.OptionButton OptOrdenPed 
            Caption         =   "Nº Pedido"
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
            TabIndex        =   177
            Top             =   550
            Width           =   1365
         End
         Begin VB.OptionButton OptOrdenArt 
            Caption         =   "Artículo"
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
            TabIndex        =   176
            Top             =   240
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   169
         Top             =   2760
         Width           =   6495
         Begin VB.TextBox txtCodigo 
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
            Index           =   68
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   156
            Top             =   735
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   68
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   171
            Text            =   "Text5"
            Top             =   735
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   67
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   155
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   67
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   170
            Text            =   "Text5"
            Top             =   360
            Width           =   3735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   15
            Left            =   390
            TabIndex        =   174
            Top             =   735
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   44
            Left            =   1080
            Top             =   735
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   14
            Left            =   390
            TabIndex        =   173
            Top             =   360
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   43
            Left            =   1080
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Artículo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   13
            Left            =   240
            TabIndex        =   172
            Top             =   120
            Width           =   810
         End
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   70
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   154
         Top             =   2400
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   69
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   153
         Top             =   2400
         Width           =   1350
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   65
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   161
         Text            =   "Text5"
         Top             =   1380
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   65
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   151
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   66
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   160
         Text            =   "Text5"
         Top             =   1755
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   66
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   152
         Top             =   1755
         Width           =   855
      End
      Begin VB.CommandButton cmdAceptarPte 
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
         Left            =   4260
         TabIndex        =   157
         Top             =   4440
         Width           =   1135
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   10
         Left            =   5460
         TabIndex        =   158
         Top             =   4440
         Width           =   1135
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   16
         Left            =   3840
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   75
         Left            =   750
         TabIndex        =   168
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   74
         Left            =   600
         TabIndex        =   167
         Top             =   2160
         Width           =   630
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   15
         Left            =   1455
         Top             =   2400
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   72
         Left            =   3150
         TabIndex        =   166
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label Label9 
         Caption         =   "Material pendiente de recibir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   19
         Left            =   600
         TabIndex        =   165
         Top             =   360
         Width           =   4455
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   41
         Left            =   1440
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   18
         Left            =   600
         TabIndex        =   164
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   17
         Left            =   750
         TabIndex        =   163
         Top             =   1380
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   42
         Left            =   1440
         Top             =   1725
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   16
         Left            =   750
         TabIndex        =   162
         Top             =   1725
         Width           =   570
      End
   End
   Begin VB.Frame FramePedidos 
      Height          =   4455
      Left            =   600
      TabIndex        =   189
      Top             =   240
      Width           =   6075
      Begin VB.TextBox txtCodigo 
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
         Index           =   76
         Left            =   1860
         MaxLength       =   15
         TabIndex        =   191
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   75
         Left            =   4260
         MaxLength       =   10
         TabIndex        =   193
         Top             =   2880
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   12
         Left            =   4470
         TabIndex        =   195
         Top             =   3720
         Width           =   1135
      End
      Begin VB.CommandButton cmdAceptarPedCom 
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
         Left            =   3210
         TabIndex        =   194
         Top             =   3720
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   74
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   192
         Top             =   2880
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   73
         Left            =   1860
         MaxLength       =   7
         TabIndex        =   190
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ped."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   600
         TabIndex        =   202
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label12 
         Caption         =   "Imprimir otros Pedidos del Proveedor:"
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
         Left            =   600
         TabIndex        =   201
         Top             =   2160
         Width           =   4065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   19
         Left            =   3900
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   5
         Left            =   840
         TabIndex        =   200
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   600
         TabIndex        =   199
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label Label12 
         Caption         =   "Informe de Pedido Compras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   570
         TabIndex        =   198
         Top             =   360
         Width           =   4335
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   18
         Left            =   1500
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   6
         Left            =   3240
         TabIndex        =   197
         Top             =   2880
         Width           =   570
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   600
         TabIndex        =   196
         Top             =   1230
         Width           =   1065
      End
   End
   Begin VB.Frame FramePedConfirma 
      Height          =   4095
      Left            =   0
      TabIndex        =   312
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtNombre 
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
         Index           =   116
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   321
         Text            =   "Text5"
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   116
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   315
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   114
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   313
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptarPedConfirma 
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
         Left            =   3690
         TabIndex        =   316
         Top             =   3240
         Width           =   1135
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   19
         Left            =   4890
         TabIndex        =   317
         Top             =   3240
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   115
         Left            =   1770
         MaxLength       =   15
         TabIndex        =   314
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   60
         Left            =   1125
         Top             =   2160
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Carta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   92
         Left            =   480
         TabIndex        =   322
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   13
         Left            =   480
         TabIndex        =   320
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "Confirmación entrega Pedido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   11
         Left            =   600
         TabIndex        =   319
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ped."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   480
         TabIndex        =   318
         Top             =   1560
         Width           =   1170
      End
   End
   Begin VB.Frame FramePasarHco 
      Height          =   4575
      Left            =   120
      TabIndex        =   97
      Top             =   120
      Width           =   6915
      Begin VB.TextBox txtNombre 
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
         Index           =   52
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   108
         Text            =   "Text5"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   52
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   100
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   51
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   103
         Text            =   "Text5"
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   51
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   99
         Top             =   2280
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   8
         Left            =   5220
         TabIndex        =   102
         Top             =   3720
         Width           =   1135
      End
      Begin VB.CommandButton cmdAceptarHco 
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
         Left            =   3990
         TabIndex        =   101
         Top             =   3720
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   50
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   98
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   29
         Left            =   1680
         Top             =   2760
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
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
         Index           =   65
         Left            =   240
         TabIndex        =   109
         Top             =   2760
         Width           =   1005
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   28
         Left            =   1680
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Operador"
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
         Index           =   64
         Left            =   240
         TabIndex        =   107
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos para el histórico: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   63
         Left            =   210
         TabIndex        =   106
         Top             =   1200
         Width           =   5490
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   14
         Left            =   2040
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Eliminación"
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
         Index           =   62
         Left            =   240
         TabIndex        =   105
         Top             =   1680
         Width           =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Albaran al histórico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   4
         Left            =   210
         TabIndex        =   104
         Top             =   480
         Width           =   4695
      End
   End
   Begin VB.Frame FrameSocios 
      Height          =   4425
      Left            =   0
      TabIndex        =   347
      Top             =   0
      Width           =   7185
      Begin VB.TextBox txtCodigo 
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
         Left            =   2190
         MaxLength       =   50
         TabIndex        =   355
         Top             =   2160
         Width           =   4275
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   354
         Top             =   1635
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   15
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   358
         Text            =   "Text5"
         Top             =   1635
         Width           =   3855
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   353
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   14
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   352
         Text            =   "Text5"
         Top             =   1260
         Width           =   3855
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   1
         Left            =   5340
         TabIndex        =   365
         Top             =   3780
         Width           =   1135
      End
      Begin VB.CommandButton cmdAceptarEtiqSocios 
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
         Left            =   4140
         TabIndex        =   363
         Top             =   3780
         Width           =   1135
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   1545
         Left            =   360
         TabIndex        =   348
         Top             =   2550
         Width           =   6255
         Begin VB.CheckBox chkMembrete 
            Caption         =   "Imprimir Membrete"
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
            Left            =   3990
            TabIndex        =   368
            Top             =   570
            Width           =   2355
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
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
            Left            =   60
            TabIndex        =   357
            Top             =   465
            Width           =   2565
         End
         Begin VB.TextBox txtNombre 
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
            Left            =   2250
            Locked          =   -1  'True
            TabIndex        =   350
            Text            =   "Text5"
            Top             =   30
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
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
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   356
            Top             =   15
            Width           =   855
         End
         Begin VB.Frame Frame15 
            Caption         =   "e-Mail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   1290
            TabIndex        =   349
            Top             =   750
            Width           =   2205
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
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
               Left            =   120
               TabIndex        =   359
               Top             =   210
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Comercial"
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
               Left            =   120
               TabIndex        =   361
               Top             =   460
               Width           =   1815
            End
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   0
            Left            =   990
            Top             =   45
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   32
            Left            =   30
            TabIndex        =   351
            Top             =   0
            Width           =   600
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "A la atención de:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   390
         TabIndex        =   367
         Top             =   2160
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   18
         Left            =   750
         TabIndex        =   366
         Top             =   1605
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   23
         Left            =   1440
         Top             =   1635
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   17
         Left            =   750
         TabIndex        =   364
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   16
         Left            =   390
         TabIndex        =   362
         Top             =   1035
         Width           =   555
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   1440
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Etiquetas a Socios"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   390
         TabIndex        =   360
         Top             =   360
         Width           =   5355
      End
   End
   Begin VB.Frame FrameEtiqProv 
      Height          =   5325
      Left            =   600
      TabIndex        =   127
      Top             =   300
      Width           =   8085
      Begin VB.TextBox txtCodigo 
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
         Index           =   62
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   89
         Top             =   3240
         Width           =   5000
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   9
         Left            =   6600
         TabIndex        =   93
         Top             =   4560
         Width           =   1135
      End
      Begin VB.CommandButton cmdAceptarEtiqProv 
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
         Left            =   5370
         TabIndex        =   92
         Top             =   4560
         Width           =   1135
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1425
         Left            =   360
         TabIndex        =   140
         Top             =   3645
         Width           =   7485
         Begin VB.Frame Frame3 
            Caption         =   "e-Mail"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   2415
            TabIndex        =   143
            Top             =   495
            Width           =   1935
            Begin VB.OptionButton OptMailAdm 
               Caption         =   "Administración"
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
               TabIndex        =   145
               Top             =   210
               Value           =   -1  'True
               Width           =   1755
            End
            Begin VB.OptionButton OptMailCom 
               Caption         =   "Compras"
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
               TabIndex        =   144
               Top             =   460
               Width           =   1755
            End
         End
         Begin VB.CheckBox chkEmail 
            Caption         =   "Enviar por e-mail"
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
            TabIndex        =   91
            Top             =   560
            Width           =   2145
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   63
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   141
            Text            =   "Text5"
            Top             =   105
            Width           =   5000
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   63
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   90
            Top             =   105
            Width           =   855
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   39
            Left            =   1080
            Top             =   105
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Carta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   9
            Left            =   240
            TabIndex        =   142
            Top             =   120
            Width           =   600
         End
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   60
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   136
         Text            =   "Text5"
         Top             =   2370
         Width           =   5000
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   60
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   87
         Top             =   2370
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   61
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   135
         Text            =   "Text5"
         Top             =   2775
         Width           =   5000
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   61
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   88
         Top             =   2775
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   59
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   86
         Top             =   1785
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   59
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "Text5"
         Top             =   1785
         Width           =   5000
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   58
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   85
         Top             =   1380
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   58
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "Text5"
         Top             =   1380
         Width           =   5000
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "A la atención de:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   8
         Left            =   600
         TabIndex        =   134
         Top             =   3240
         Width           =   1785
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CPostal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   600
         TabIndex        =   139
         Top             =   2130
         Width           =   825
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   37
         Left            =   1440
         Top             =   2370
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   6
         Left            =   780
         TabIndex        =   138
         Top             =   2370
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   38
         Left            =   1440
         Top             =   2775
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   7
         Left            =   780
         TabIndex        =   137
         Top             =   2715
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   4
         Left            =   780
         TabIndex        =   133
         Top             =   1725
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   36
         Left            =   1440
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   3
         Left            =   780
         TabIndex        =   132
         Top             =   1380
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   600
         TabIndex        =   131
         Top             =   1035
         Width           =   1125
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   35
         Left            =   1440
         Top             =   1380
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Etiquetas Proveedores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   1
         Left            =   600
         TabIndex        =   130
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame FrameClientes 
      Height          =   6015
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   8955
      Begin VB.TextBox txtCodigo 
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
         Index           =   41
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   46
         Top             =   4695
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   42
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   47
         Top             =   5070
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   41
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   4695
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   42
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   5070
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   38
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   3270
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   37
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   2895
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   38
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   43
         Top             =   3270
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   37
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   42
         Top             =   2895
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   6
         Left            =   7590
         TabIndex        =   49
         Top             =   5550
         Width           =   1135
      End
      Begin VB.CommandButton cmdAceptarClien 
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
         Left            =   6330
         TabIndex        =   48
         Top             =   5550
         Width           =   1135
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   38
         Top             =   1140
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   39
         Top             =   1515
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   33
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text5"
         Top             =   1140
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   34
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text5"
         Top             =   1515
         Width           =   3645
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   40
         Top             =   2010
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   41
         Top             =   2385
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   35
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text5"
         Top             =   2010
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   36
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text5"
         Top             =   2385
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   39
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   44
         Top             =   3795
         Width           =   615
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   40
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   45
         Top             =   4170
         Width           =   615
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   39
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text5"
         Top             =   3795
         Width           =   3645
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   40
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "Text5"
         Top             =   4170
         Width           =   3645
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   510
         Left            =   8160
         Picture         =   "frmListadoOfer.frx":05A2
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2505
         Width           =   435
      End
      Begin VB.CommandButton cmdSubir 
         Height          =   510
         Left            =   8160
         Picture         =   "frmListadoOfer.frx":08AC
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1720
         Width           =   435
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   6480
         TabIndex        =   56
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   2
         Left            =   900
         TabIndex        =   75
         Top             =   4695
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   1
         Left            =   900
         TabIndex        =   74
         Top             =   5070
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Situación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   45
         Left            =   600
         TabIndex        =   73
         Top             =   4440
         Width           =   975
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   21
         Left            =   1635
         Top             =   4695
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   22
         Left            =   1635
         Top             =   5085
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   18
         Left            =   1635
         Top             =   3300
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1635
         Top             =   2895
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   51
         Left            =   600
         TabIndex        =   70
         Top             =   2655
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   48
         Left            =   900
         TabIndex        =   69
         Top             =   3270
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   50
         Left            =   900
         TabIndex        =   68
         Top             =   2895
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   55
         Left            =   900
         TabIndex        =   67
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   56
         Left            =   900
         TabIndex        =   66
         Top             =   1515
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Actividad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   49
         Left            =   600
         TabIndex        =   65
         Top             =   900
         Width           =   990
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   13
         Left            =   1635
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1635
         Top             =   1530
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Informe de Clientes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   64
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   57
         Left            =   900
         TabIndex        =   63
         Top             =   2010
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   58
         Left            =   900
         TabIndex        =   62
         Top             =   2385
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   48
         Left            =   600
         TabIndex        =   61
         Top             =   1770
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1635
         Top             =   2010
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1635
         Top             =   2415
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Desde"
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
         Index           =   59
         Left            =   900
         TabIndex        =   60
         Top             =   3795
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta"
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
         Index           =   60
         Left            =   900
         TabIndex        =   59
         Top             =   4170
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   47
         Left            =   600
         TabIndex        =   58
         Top             =   3540
         Width           =   765
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   19
         Left            =   1635
         Top             =   3795
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   20
         Left            =   1635
         Top             =   4185
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   46
         Left            =   6480
         TabIndex        =   57
         Top             =   1200
         Width           =   1980
      End
   End
   Begin VB.Frame FrameCompras 
      Height          =   5205
      Left            =   360
      TabIndex        =   233
      Top             =   480
      Width           =   7035
      Begin VB.CheckBox chkDatosAlbaranes 
         Caption         =   "Datos albaranes"
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
         Left            =   2760
         TabIndex        =   242
         Top             =   3960
         Width           =   2205
      End
      Begin VB.Frame Frame9 
         Caption         =   "Agrupar por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   945
         Left            =   360
         TabIndex        =   260
         Top             =   3880
         Width           =   2325
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia"
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
            TabIndex        =   240
            Top             =   270
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptCompras 
            Caption         =   "Familia, Artículo"
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
            Left            =   240
            TabIndex        =   241
            Top             =   585
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   16
         Left            =   5640
         TabIndex        =   244
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarCompras 
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
         Left            =   4560
         TabIndex        =   243
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   91
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   235
         Top             =   1665
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   91
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   246
         Text            =   "Text5"
         Top             =   1665
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   90
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   234
         Top             =   1260
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
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
         Index           =   90
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   245
         Text            =   "Text5"
         Top             =   1260
         Width           =   3975
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   92
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   236
         Top             =   2280
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   93
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   237
         Top             =   2280
         Width           =   1350
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   254
         Top             =   2580
         Width           =   6495
         Begin VB.TextBox txtNombre 
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
            Index           =   94
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   256
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   94
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   238
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
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
            Index           =   95
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   255
            Text            =   "Text5"
            Top             =   765
            Width           =   3855
         End
         Begin VB.TextBox txtCodigo 
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
            Index           =   95
            Left            =   1380
            MaxLength       =   16
            TabIndex        =   239
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Familia"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   20
            Left            =   240
            TabIndex        =   259
            Top             =   120
            Width           =   780
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   50
            Left            =   1080
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
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
            Index           =   12
            Left            =   390
            TabIndex        =   258
            Top             =   360
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   51
            Left            =   1080
            Top             =   765
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
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
            Index           =   11
            Left            =   390
            TabIndex        =   257
            Top             =   765
            Width           =   570
         End
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   24
         Left            =   750
         TabIndex        =   253
         Top             =   1665
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   49
         Left            =   1440
         Top             =   1665
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   23
         Left            =   750
         TabIndex        =   252
         Top             =   1260
         Width           =   600
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   22
         Left            =   600
         TabIndex        =   251
         Top             =   915
         Width           =   1125
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   48
         Left            =   1440
         Top             =   1260
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Compras por Familia/Artículo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Index           =   21
         Left            =   600
         TabIndex        =   250
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Index           =   88
         Left            =   3210
         TabIndex        =   249
         Top             =   2280
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   27
         Left            =   1455
         Top             =   2280
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   87
         Left            =   600
         TabIndex        =   248
         Top             =   2010
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Index           =   83
         Left            =   750
         TabIndex        =   247
         Top             =   2280
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   28
         Left            =   3840
         Top             =   2280
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmListadoOfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
    '(ver opciones en frmListado)
        
        
        
    '315:  Envio por mail de las facturas
        
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta/pedido a imprimir

Public CodClien As String 'Para seleccionar inicialmente las ofertas del Cliente
                          'en el listado de Recordatorio de Ofertas y de Valoracion de Ofertas

Public FecEntre As String 'Para pasar inicialmente la fecha de entrega de la Oferta que se va a pasar a pedido
                          'como la fecha de entega del PEdido
                          
Private NomTabla As String
Private NomTablaLin As String

Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Attribute frmMtoCartasOfe.VB_VarHelpID = -1
Private WithEvents frmMtoCliente As frmFacClientes
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoProve As frmComProveedores
Attribute frmMtoProve.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoActiv As frmFacActividades
Attribute frmMtoActiv.VB_VarHelpID = -1
Private WithEvents frmMtoSitua As frmFacSituaciones
Attribute frmMtoSitua.VB_VarHelpID = -1
Private WithEvents frmMtoIncid As frmIncidencias
Attribute frmMtoIncid.VB_VarHelpID = -1
Private WithEvents frmMtoArtic As frmAlmArticulos
Attribute frmMtoArtic.VB_VarHelpID = -1
Private WithEvents frmMtoFamilia As frmAlmFamiliaArticulo
Attribute frmMtoFamilia.VB_VarHelpID = -1
Private WithEvents frmSocio As frmGesSocios
Attribute frmSocio.VB_VarHelpID = -1


'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'codigo postal
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmMen As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen.VB_VarHelpID = -1
Private WithEvents frmMen2 As frmMensajes  'Form Mensajes para mostrar las etiquetas a imprimir
Attribute frmMen2.VB_VarHelpID = -1



'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'cadena con los parametros q se pasan a Crystal R.
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'nuevo Febrero 2010
Private cadPDFrpt As String 'Nombre del informe a enviar por email
Private vImprimedirecto As Boolean '
Private CadenaParaEnvioMail As String
'-------------------------------------



Dim indCodigo As Byte 'indice para txtCodigo
Dim Codigo As String 'Código para FormulaSelection de Crystal Report

Dim PrimeraVez As Boolean

Dim kCampo As Integer



Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub





Private Sub cboTipomov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkConfirmPed_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDatosAlbaranes_Click(Index As Integer)
    If Index = 0 Then
        Label4(90).Caption = "Fecha"
        If Me.chkDatosAlbaranes(0).Value = 1 Then Label4(90).Caption = Label4(90).Caption & " albaran"
    Else
        Label4(87).Caption = "Fecha"
        If Me.chkDatosAlbaranes(1).Value = 1 Then Label4(87).Caption = Label4(87).Caption & " albaran"
    End If
End Sub

Private Sub chkDatosAlbaranes_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkDetallaArticulo_Click()
    Me.FrameDetallaArticulo.visible = chkDetallaArticulo.Value = 1
End Sub

Private Sub chkDetallaArticulo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkEmail_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpress KeyAscii
End Sub



Private Sub chkImpSaldo_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkMail_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub chkPendientes_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub cmdAceptarEtiqSocios_Click()
'305: Listado para etiquetas de proveedor
'306: Listado para cartas a proveedor
Dim campo As String

    InicializarVbles
    
    'si es listado de CARTAS/eMAIL a proveedores comprobar que se ha seleccionado
    'una carta para imprimir
    If OpcionListado = 191 Then
        If txtcodigo(5).Text = "" Then
            MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
            Exit Sub
        End If
        
        'Parametro cod. carta
        cadParam = "|pCodCarta= " & txtcodigo(5).Text & "|"
        numParam = numParam + 1
        
        'Nombre fichero .rpt a Imprimir
        If Me.chkMembrete.Value = 1 Then
            nomRPT = "rFacSocioCartaLog.rpt" '"rComProveCarta.rpt"
        Else
            nomRPT = "rFacSocioCarta.rpt" '"rComProveCarta.rpt"
        End If
        Titulo = "Cartas a Socios"
        conSubRPT = True
        
    Else 'ETIQUETAS
        cadParam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacSocioEtiq.rpt" '"rComProveEtiq.rpt"
        Titulo = "Etiquetas de Socios"
        conSubRPT = False
    End If
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H SOCIOS
    '--------------------------------------------
     If txtcodigo(14).Text <> "" Or txtcodigo(15).Text <> "" Then
        campo = "{sclien.codclien}"
        'Parametro Desde/Hasta Proveedor
        If Not PonerDesdeHasta(campo, "N", 14, 15, "") Then Exit Sub
    End If
    
    
    ' Añadimos la condicion de que tengan nro de uve
    If Not AnyadirAFormula(cadFormula, "{sclien.numeruve} <> 0 and not isnull({sclien.numeruve})") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{sclien.numeruve} <> 0 and not {sclien.numeruve} is null") Then Exit Sub
    
    
    ' Añadimos la condicion de que no tengan fecha de baja
    If Not AnyadirAFormula(cadFormula, "isnull({sclien.fechabaj})") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{sclien.fechabaj} is null") Then Exit Sub
    
    '====================================================
        
    'Sacamos que situaciones de los socios quieren
    Set frmMen2 = New frmMensajes
    frmMen2.OpcionMensaje = 24 'Situaciones de socios
    frmMen2.Show vbModal
    Set frmMen2 = Nothing
    If cadSelect = "" Then Exit Sub
        
        
    'Parametro a la Atencion de
    cadParam = cadParam & "pAtencion=""" & txtcodigo(7).Text & """|"
    numParam = numParam + 1
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme("sclien", cadSelect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadWHERE = cadSelect
    frmMen.OpcionMensaje = 23 'Etiquetas socios
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    If OpcionListado = 191 And Me.chkEMAIL(2).Value = 1 Then
        'Enviarlo por e-mail
        nomRPT = "rFacSocioCartaLog.rpt" '"rComProveCarta.rpt"
        
        EnviarEMailMulti cadSelect, Titulo, "rFacSociosCarta.rpt", "sclien" 'email para socios
    Else
        LlamarImprimir False, False
    End If
    

End Sub

Private Sub cmdAceptarAlbCom_Click()
'Solicitar datos para Generar Albaran  a partir de Pedido de Compras
Dim Cad As String

    Cad = txtcodigo(47).Text & "|"
    Cad = Cad & txtcodigo(48).Text & "|"
    Cad = Cad & txtcodigo(49).Text & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub cmdAceptarCierre_Click()
'Cierre de caja del TPV
Dim campo As String
Dim devuelve As String


    InicializarVbles
    
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1


    'comprobar que se ha introducido FECHA
    '---------------------------------------------------------
    If Trim(txtcodigo(88).Text) <> "" Or Trim(txtcodigo(89).Text) <> "" Then
        'Para Crystal Report
        campo = "{scafac.fecfactu}"
        devuelve = "pDHFecha=""FECHA: " 'Parametro Desde/Hasta Fecha
        If Not PonerDesdeHasta(campo, "F", 88, 89, devuelve) Then Exit Sub
    Else
        MsgBox "Debe introducir la fecha de cierre.    ", vbExclamation
        Exit Sub
    End If
    
    
    '---- Seleccionar solo las facturas que vienen de TICKET del TPV
    campo = "{scafac1.numventa}"
    campo = "(NOT ISNULL(" & campo & ")) and (" & campo & "<>0)"
    If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    
    campo = "{scafac1.numtermi} >0"
    If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
    If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    
    '---- seleccionar solo el tipo pago: 0=efectivo,2=talon, 3=pagare, 6=tarjeta credito
    campo = "{sforpa.tipforpa} in [0,2,3,6]"
    If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
    campo = "{sforpa.tipforpa} in (0,2,3,6)"
    If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    
    ' ---- [14/10/2009] [LAURA]: para q muestre los ATI y tambien los ALV generados en el TPV
'    'Seleccionar solo las facturas que los albaranes fueron generados en el TPV
'    'para ello seleccionar que scafac1.codtipoa='ATI'
'    campo = "{scafac1.codtipoa} = 'ATI'"
'    If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    ' ----
    
    'ver si hay registros seleccionados para mostrar en el informe
    campo = "(scafac INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom and scafac.numfactu=scafac1.numfactu and scafac.fecfactu=scafac1.fecfactu)  INNER JOIN sforpa ON scafac.codforpa = sforpa.codforpa "
    If Not HayRegParaInforme(campo, cadSelect) Then Exit Sub
    
    Titulo = "Cierre de Caja"
    If Me.optForpago(0).Value = True Then
        'informe por Forma de Pago
        nomRPT = "rTPVcierreFP.rpt"
    Else
        'informe por Tipo de Forma de Pago
        nomRPT = "rTPVcierre.rpt"
    End If
    conSubRPT = True
    LlamarImprimir False, False
     
End Sub

Private Sub cmdAceptarClien_Click()
'Listado de Clientes
Dim campo As String, devuelve As String
Dim numOp As Byte

    InicializarVbles
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H ACTIVIDAD
    '--------------------------------------------
     If txtcodigo(33).Text <> "" Or txtcodigo(34).Text <> "" Then
        campo = "{scliente.codactiv}"
        'Parametro Desde/Hasta Actividad
        devuelve = "pDHActividad=""Actividad: "
        If Not PonerDesdeHasta(campo, "N", 33, 34, devuelve) Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H AGENTE
    '--------------------------------------------
     If txtcodigo(39).Text <> "" Or txtcodigo(40).Text <> "" Then
        campo = "{scliente.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 39, 40, devuelve) Then Exit Sub
    End If
    
    'Cadena para seleccion D/H SITUACION
    '--------------------------------------------
     If txtcodigo(41).Text <> "" Or txtcodigo(42).Text <> "" Then
        campo = "{scliente.codsitua}"
        'Parametro Desde/Hasta Situacion
        devuelve = "pDHSituacion=""Situación: "
        If Not PonerDesdeHasta(campo, "N", 41, 42, devuelve) Then Exit Sub
    End If
    
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
'    numOp = PonerGrupo(3, ListView1.ListItems(3).Text)
'    numOp = PonerGrupo(4, ListView1.ListItems(4).Text)

    cadSelect = cadFormula
    If Not HayRegParaInforme("scliente", cadSelect) Then Exit Sub
     
    LlamarImprimir False, False
End Sub


Private Sub cmdAceptarClienInac_Click()
'46: Informe de Clientes Inactivos
'47: Informe de Altas Nuevos Clientes
'90: Informe Etiquetas de clientes
Dim campo As String, devuelve As String

    InicializarVbles
    
    If OpcionListado = 46 Then
        'Comprobar que se ha introdicido una FECHA de Inactividad
        If txtcodigo(31).Text = "" Then
            MsgBox "Debe introducir la Fecha de Inactividad para el informe.", vbInformation
            Exit Sub
        End If
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacClienInactivos.rpt"
        Titulo = "Clientes Inactivos"
        
    ElseIf OpcionListado = 48 Then
        'Comprobar si se ha introducido D/H FECHA Alta
        If txtcodigo(31).Text = "" And txtcodigo(32).Text = "" Then
            MsgBox "Debe introducir algún intervalo de Fechas de Alta.", vbInformation
            Exit Sub
        End If
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rFacClienAltas.rpt"
    End If
    
    '===================================================
    '============ PARAMETROS ===========================
    'Nombre de la Empresa
    cadParam = "|pEmpresa=""" & vParam.NombreEmpresa & """|"
    numParam = numParam + 1
    
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtcodigo(27).Text <> "" Or txtcodigo(28).Text <> "" Then
        campo = "{scliente.codclien}"
        'Parametro Desde/Hasta Cliente
        devuelve = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 27, 28, devuelve) Then Exit Sub
    End If
    
    
    'Cadena para seleccion D/H AGENTE
    '--------------------------------------------
     If txtcodigo(29).Text <> "" Or txtcodigo(30).Text <> "" Then
        campo = "{scliente.codagent}"
        'Parametro Desde/Hasta Agente
        devuelve = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(campo, "N", 29, 30, devuelve) Then Exit Sub
    End If
    
    
    
    If OpcionListado = 90 Or OpcionListado = 91 Then '90: Etiquetas de clientes
                                                     '91: Cartas a clientes
        'Cadena para seleccion D/H ACTIVIDAD
        '--------------------------------------------
         If txtcodigo(53).Text <> "" Or txtcodigo(54).Text <> "" Then
            campo = "{scliente.codactiv}"
            'Parametro Desde/Hasta Actividad
            devuelve = "pDHActividad=""Actividad: "
            If Not PonerDesdeHasta(campo, "N", 53, 54, devuelve) Then Exit Sub
        End If
                        
        'Cadena para seleccion D/H COD. POSTAL
        '--------------------------------------------
         If txtcodigo(55).Text <> "" Or txtcodigo(56).Text <> "" Then
            campo = "{scliente.codpobla}"
            'Parametro Desde/Hasta cod. Postal
            devuelve = "pDHcpostal=""CPostal: "
            If Not PonerDesdeHasta(campo, "T", 55, 56, devuelve) Then Exit Sub
        End If
        
        'Cadena para seleccion SITUACION
        '--------------------------------------------
        If txtcodigo(57).Text <> "" Then
            campo = "{scliente.codsitua}=" & txtcodigo(57).Text
            If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        End If
        
        
        'ENERO 2010
        'Si no tiene  la marca de correo NO puede seleccionar cliente
        campo = "{scliente.enviocorreo}=1"
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
        
        'Parametro a la Atencion de
        cadParam = cadParam & "pAtencion=""Att. " & txtcodigo(0).Text & """|"
        numParam = numParam + 1
        
        'seleccionamos todos los clientes por defecto,
        'pero si seleccionamos clientes con mantenimientos o sin mantenimientos
         'Comprobar si hay registros a Mostrar antes de abrir el Informe
        cadSelect = QuitarCaracterACadena(cadFormula, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
        
        devuelve = ""
        If Me.OptCliConMante Then
            devuelve = ListaClientesMante(cadSelect)
            If devuelve <> "" Then
                cadFormula = "{scliente.codclien} IN [" & devuelve & "]"
                cadSelect = "scliente.codclien IN (" & devuelve & ")"
            End If
        ElseIf Me.OptCliSinMante Then
            devuelve = ListaClientesMante(cadSelect)
            If devuelve <> "" Then
                campo = " NOT( {scliente.codclien}  IN [" & devuelve & "])"
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                campo = " scliente.codclien NOT IN (" & devuelve & ")"
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
        End If
        
        If OpcionListado = 90 Then
            
            devuelve = ListaClientesDesdeHastaFactura2()
            'Puede haber puesto desde hasta datos factura
            If devuelve <> "" Then
                campo = " ( {scliente.codclien}  IN [" & devuelve & "])"
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
        End If
        
        If OpcionListado = 90 Then 'Etiquetas
            'Nombre fichero .rpt a Imprimir
            
            'NUEVO. Igual deberiamos utilizar la clase: CParamTPV
            nomRPT = DevuelveDesdeBD(conAri, "documrpt", "scryst", "codcryst", "20")
            If nomRPT = "" Then nomRPT = "rFacClienEtiq.rpt"
            Titulo = "Etiquetas de Clientes"
            conSubRPT = False
        Else '91: CARTA/e-MAil
            'Parametro cod. carta
            cadParam = "|pCodCarta= " & txtcodigo(64).Text & "|"
            numParam = numParam + 1
            
            'Nombre fichero .rpt a Imprimir
            nomRPT = "rFacClienCarta.rpt"
            Titulo = "Cartas a Clientes"
            conSubRPT = True
        End If
    Else
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        cadSelect = QuitarCaracterACadena(cadFormula, "{")
        cadSelect = QuitarCaracterACadena(cadSelect, "}")
    End If
    
    If OpcionListado = 46 Then
        'Seleccionar aquellos cliente que campo sclien.fechamov < fecha Inactividad
        If txtcodigo(31).Text <> "" Then
            campo = "scliente.fechamov"
            devuelve = txtcodigo(31).Text
            devuelve = "(isnull({scliente.fechamov}) or {" & campo & "} < Date(" & Year(devuelve) & ", " & Month(devuelve) & ", " & Day(devuelve) & "))"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            devuelve = "(" & campo & " < '" & Format(txtcodigo(31).Text, FormatoFecha) & "' OR isnull(scliente.fechamov))"
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
            devuelve = "pFechaMov=""" & txtcodigo(31).Text & """|"
            cadParam = cadParam & devuelve
            numParam = numParam + 1
        End If
        
    ElseIf OpcionListado = 48 Then
        'Cadena para seleccion D/H FECHA
        '--------------------------------------------
        If txtcodigo(31).Text <> "" Or txtcodigo(32).Text <> "" Then
          'Para Crystal Report
            campo = "{scliente.fechaalt}"
            'Parametro Desde/Hasta Fecha
            devuelve = "pDHFecha=""Fecha Alta: "
            If Not PonerDesdeHasta(campo, "F", 31, 32, devuelve) Then Exit Sub
        End If
    End If
        
    If Not HayRegParaInforme("scliente", cadSelect) Then Exit Sub
    
    If OpcionListado = 90 Or OpcionListado = 91 Then 'OpcionListado = 90 'Etiquetas clientes
        Set frmMen = New frmMensajes
        frmMen.cadWHERE = cadSelect
        frmMen.OpcionMensaje = 8 'Etiquetas clientes
        frmMen.Show vbModal
        Set frmMen = Nothing
        If cadSelect = "" Then Exit Sub
        
        If OpcionListado = 91 And Me.chkEMAIL(1).Value = 1 Then
            'Enviarlo por e-mail
            EnviarEMailMulti cadSelect, Titulo, "rFacClienCarta.rpt", "sclien" 'email para clientes
        Else
            LlamarImprimir False, False
        End If
    Else
        LlamarImprimir False, False
    End If
    
End Sub


Private Sub cmdAceptarCompras_Click()
'Listados de Compras
Dim campo As String
Dim Cad As String
Dim Tabla As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtcodigo(90).Text <> "" Or txtcodigo(91).Text <> "" Then
        campo = "{scafpc.codprove}"
        'Parametro Desde/Hasta Proveedor
        Cad = "pDHProve=""Proveedor: "
        If Not PonerDesdeHasta(campo, "N", 90, 91, Cad) Then Exit Sub
    End If
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(92).Text <> "" Or txtcodigo(93).Text <> "" Then
        'Para fam/articulo con albaranaes
        If OpcionListado = 311 And Me.chkDatosAlbaranes(1).Value = 1 Then
            campo = "{scafpa.fechaalb}"
        Else
            campo = "{scafpc.fecfactu}"
        End If
        Cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 92, 93, Cad) Then Exit Sub
    End If
    
    Tabla = "scafpc"
    If OpcionListado = 311 Then
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtcodigo(94).Text <> "" Or txtcodigo(95).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            Cad = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 94, 95, Cad) Then Exit Sub
            
            
            If Me.chkDatosAlbaranes(1).Value = 0 Then
                Tabla = "( scafpc INNER JOIN slifpc ON scafpc.codprove=slifpc.codprove AND scafpc.numfactu=slifpc.numfactu "
                Tabla = Tabla & " AND scafpc.fecfactu=slifpc.fecfactu )"
                Tabla = Tabla & " INNER JOIN sartic ON slifpc.codartic=sartic.codartic "
        
        
            Else
                
            
            
            
            End If
        
        End If
    End If
        
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    If OpcionListado = 312 Then
        'en esta opcion ver si hay albaranes
        cadSelect = Replace(cadSelect, Tabla, "scafpa")
        cadSelect = Replace(cadSelect, "fecfactu", "fechaalb")
        Tabla = "scafpa"
    End If
    
    'Para fam/articulo con albaranaes
    If OpcionListado = 311 And Me.chkDatosAlbaranes(1).Value = 1 Then
        'Es un contador de un UNION.
        'Hay que comprobar si hay reg en factuaras Y albaranes
        If Not ContadorDelUnion(False) Then
            MsgBox "No existen valores con esos parametros", vbExclamation
            Exit Sub
        End If
    
    Else
        If Not HayRegParaInforme(Tabla, cadSelect) Then
            If OpcionListado <> 312 Then Exit Sub
        
            If Not HayRegParaInforme("scaalp", cadSelect) Then Exit Sub
        End If
    End If
    
    If OpcionListado = 312 Then
    
    
        'insertar en la tmpInformes
        BorrarTempInformes
        
        'en esta opcion ver si hay albaranes
        Cad = Replace(cadSelect, Tabla, "scaalp")
        Cad = Replace(Cad, "fecfactu", "fechaalb")
        
        'insertar los albaranes q cumple la seleccion
        If Not CargarTmpInformes_Compras_312("scaalp", Cad) Then Exit Sub
        
        
        'insertar los albaranes de facturas q cumple la seleccion
        If Not CargarTmpInformes_Compras_312(Tabla, cadSelect) Then Exit Sub
        
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
    End If
    
    
    
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    conSubRPT = False
    If OpcionListado = 311 Then
        If Me.OptCompras(0).Value = True Then
            nomRPT = "rComEstProFam"
            Titulo = "Listado Compras por Familia"
            conSubRPT = True
        Else
            nomRPT = "rComEstProArt"
            Titulo = "Listado Compras por Artículo"
        End If
        
        If Me.chkDatosAlbaranes(1).Value = 1 Then
            nomRPT = nomRPT & "alb"
            Titulo = Titulo & " (con albaranes)"
            
            
            'Cambiamos los desde hasta
            'En la cadena selleccion cambiamos las tabla por
            cadFormula = Replace(cadFormula, "scafpa", "Command")
            cadFormula = Replace(cadFormula, "scafpc", "Command")
            cadFormula = Replace(cadFormula, "sartic", "Command")
            cadFormula = Replace(cadFormula, "slifpc", "Command")
            
            
            
        End If
        nomRPT = nomRPT & ".rpt"
        
        
    ElseIf OpcionListado = 310 Then
        nomRPT = "rComEstProImp.rpt"
        Titulo = "Listado Compras por Proveedor"
    Else '312: Albaranes x porveedor
        nomRPT = "rComEstProAlb.rpt"
        Titulo = "Listado albaranes por Proveedor"
    End If
    
    
    LlamarImprimir False, False
    
    If OpcionListado = 312 Then BorrarTempInformes
End Sub

Private Sub CmdAceptarEstCliente_Click()
'Listados estadistica ventas por familia
'Listados de Compras
Dim campo As String
Dim Cad As String
Dim Tabla As String


    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtcodigo(3).Text <> "" Or txtcodigo(4).Text <> "" Then
        campo = "{scafaccli.codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 3, 4, Cad) Then Exit Sub
    End If
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    'MOdificacion  18 Novi 2008
    'Las estadisticas son sobre facturas.... Y ALBARANES!!!!
    'La fecha no se la puedo pasar porque en el union hacer referencia a dos campos
    'fecfactu(factura) y fechaalb (albaranes)
    'para ello hay un parametro en el informe
  
    If txtcodigo(1).Text <> "" Or txtcodigo(2).Text <> "" Then
        campo = "{scafaccli.fecfactu}"
        Cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 1, 2, Cad) Then Exit Sub
    End If
    
    Tabla = "scafaccli"

    ' [Monica]04/04/2011: Añadido el combo de tipo de factura de venta socio
    If Mid(Combo2.Text, 1, 5) <> "Todos" Then
        If Not AnyadirAFormula(cadFormula, "{scafaccli.codtipom} = '" & Mid(Combo2.Text, 1, 3) & "'") Then Exit Sub
        If Not AnyadirAFormula(cadSelect, "{scafaccli.codtipom} = '" & Mid(Combo2.Text, 1, 3) & "'") Then Exit Sub
    End If
        
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    If Me.chkDatosAlbaranes(0).Value = 0 Or Me.OptPorFamilia.Value = True Then
        If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    Else
        'Es un contador de un UNION
        If Not ContadorDelUnion(True) Then
            MsgBox "No existen valores con esos parametros", vbExclamation
            Exit Sub
        End If
    End If
    
    
    nomRPT = "rFacEstClienteImp.rpt"
    Titulo = "Detalle Facturación Clientes"
    conSubRPT = False
    
    LlamarImprimir False, False
    

End Sub

Private Sub cmdAceptarEstVentas_Click()
'Listados estadistica ventas por familia
'Listados de Compras
Dim campo As String
Dim Cad As String
Dim Tabla As String


    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = "|pNomEmpre=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H CLIENTE
    '--------------------------------------------
     If txtcodigo(96).Text <> "" Or txtcodigo(97).Text <> "" Then
        campo = "{scafac.codclien}"
        'Parametro Desde/Hasta Cliente
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(campo, "N", 96, 97, Cad) Then Exit Sub
    End If
   
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    'MOdificacion  18 Novi 2008
    'Las estadisticas son sobre facturas.... Y ALBARANES!!!!
    'La fecha no se la puedo pasar porque en el union hacer referencia a dos campos
    'fecfactu(factura) y fechaalb (albaranes)
    'para ello hay un parametro en el informe
  
    If txtcodigo(98).Text <> "" Or txtcodigo(99).Text <> "" Then
        If Me.chkDatosAlbaranes(0).Value = 1 And Me.OptPorFamilia.Value = False Then
            campo = "{scafac1.fechaalb}"
        Else
            campo = "{scafac.fecfactu}"
        End If
        Cad = "pDHFecha=""Fecha Fact.: "
        If Not PonerDesdeHasta(campo, "F", 98, 99, Cad) Then Exit Sub
    End If
    
    Tabla = "scafac"

    If OpcionListado = 230 Then
        campo = ""  'Para comprobar que alguno de los campos es distinto de ""
        
        
        '---------------   VENTAS x FAMILIA / ARTICULO
        If Me.chkDetallaArticulo.Value = 1 Then
            If txtcodigo(112).Text <> "" Or txtcodigo(112).Text <> "" Then
                campo = "{slifac.codArtic}"
                Cad = "pDHFamilia=""Artículo: "
                If Not PonerDesdeHasta(campo, "T", 112, 113, Cad) Then Exit Sub
            End If
        End If
    
    
        'Cadena para seleccion D/H FAMILIA
        '--------------------------------------------
         If txtcodigo(100).Text <> "" Or txtcodigo(101).Text <> "" Then
            campo = "{sartic.codfamia}"
            'Parametro Desde/Hasta Familia
            Cad = "pDHFamilia=""Familia: "
            If Not PonerDesdeHasta(campo, "N", 100, 101, Cad) Then Exit Sub
        End If
        
        
        'Si por algun campo de los de arriba es <>"" entonces tenemos que meter esto
        If campo <> "" Then
            If Me.chkDatosAlbaranes(0).Value = 0 Then
                'Sin albaranes
                Tabla = "( scafac INNER JOIN slifac ON scafac.codtipom=slifac.codtipom AND scafac.numfactu=slifac.numfactu "
                Tabla = Tabla & " AND scafac.fecfactu=slifac.fecfactu )"
                Tabla = Tabla & " INNER JOIN sartic ON slifac.codartic=sartic.codartic "
            End If
        End If
    Else
        ' [Monica]04/04/2011: Añadido el combo de tipo de factura de venta socio
        If Mid(Combo1.Text, 1, 5) <> "Todos" Then
            If Not AnyadirAFormula(cadFormula, "{scafac.codtipom} = '" & Mid(Combo1.Text, 1, 3) & "'") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{scafac.codtipom} = '" & Mid(Combo1.Text, 1, 3) & "'") Then Exit Sub
        End If
        
    End If
    
    'Comprobar si hay registros para mostrar en el informe
    '========================================================
    If Me.chkDatosAlbaranes(0).Value = 0 Or Me.OptPorFamilia.Value = True Then
        If Not HayRegParaInforme(Tabla, cadSelect) Then Exit Sub
    Else
        'Es un contador de un UNION
        If Not ContadorDelUnion(True) Then
            MsgBox "No existen valores con esos parametros", vbExclamation
            Exit Sub
        End If
    End If
    
    
    'Abrir el listado
    '=======================================
    'Nombre fichero .rpt a Imprimir
    If OpcionListado = 230 Then
    
        If Me.OptPorCliente.Value = True Then 'Agrupar por cliente y familia
            If Me.chkDetallaArticulo.Value = 0 Then
                nomRPT = "rFacEstCliFam"
                Titulo = "Listado Ventas por cliente/familia"
                conSubRPT = True
            Else
                nomRPT = "rFacEstCliFamArt"
                Titulo = "Listado ventas por cliente/familia/artículo"
                conSubRPT = False
            End If
            
            
            If Me.chkDatosAlbaranes(0).Value = 1 Then
                nomRPT = nomRPT & "Alb"
                Titulo = Titulo & "(Con albaranes)"
                
                'En la cadena seleccion cambiamos las tabla por
                cadFormula = Replace(cadFormula, "scafac1", "Command")
                cadFormula = Replace(cadFormula, "scafac", "Command")
                cadFormula = Replace(cadFormula, "sartic", "Command")
                cadFormula = Replace(cadFormula, "slifac", "Command")
            End If
        ' ---- [09/11/2009] [LAURA] : agrupar por cliente/familia o solo por familia
        '                             en ambos casos puede detallar articulo
        ElseIf Me.OptPorFamilia.Value = True Then 'agrupar solo por familia
            If Me.chkDetallaArticulo.Value = 0 Then
                nomRPT = "rFacEstFam"
                Titulo = "Listado Ventas por familia"
                conSubRPT = True
            Else
                nomRPT = "rFacEstFamArt"
                Titulo = "Listado ventas por familia/artículo"
                conSubRPT = False
            End If
            
        End If
        
        
        nomRPT = nomRPT & ".rpt"
    Else
        nomRPT = "rFacEstCliImp.rpt"
        Titulo = "Detalle Facturación Socios"
        conSubRPT = False
    End If
    
    
    LlamarImprimir False, False
    
End Sub

Private Function ContadorDelUnion(Compras As Boolean) As Boolean
Dim C As String

    'Con albaranes
    Codigo = cadSelect
    Codigo = QuitarCaracterACadena(Codigo, "{")
    Codigo = QuitarCaracterACadena(Codigo, "}")
    
    
    ContadorDelUnion = False
    If Compras Then
            C = "(SELECT count(*) FROM   (((`scafac1` `scafac1` INNER JOIN `scafac` `scafac` ON"
            C = C & " ((`scafac1`.`codtipom`=`scafac`.`codtipom`) AND (`scafac1`.`numfactu`=`scafac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`scafac`.`fecfactu`)) INNER JOIN `slifac` `slifac` ON"
            C = C & " ((((`scafac1`.`codtipom`=`slifac`.`codtipom`) AND (`scafac1`.`numfactu`=`slifac`.`numfactu`))"
            C = C & " AND (`scafac1`.`fecfactu`=`slifac`.`fecfactu`)) AND (`scafac1`.`numalbar`=`slifac`.`numalbar`))"
            C = C & " AND (`scafac1`.`codtipoa`=`slifac`.`codtipoa`)) INNER JOIN `sartic` `sartic`"
            C = C & " ON `slifac`.`codartic`=`sartic`.`codartic`) INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
            
            If Codigo <> "" Then C = C & " WHERE " & Codigo
            C = C & ") + ("
            C = C & " SELECT count(*) from ((`slialb` INNER JOIN scaalb ON ((`slialb`.`codtipom`=`scaalb`.`codtipom`) AND"
            C = C & " (`slialb`.`numalbar`=`scaalb`.`numalbar`)))"
            C = C & " INNER JOIN `sartic` `sartic` ON `slialb`.`codartic`=`sartic`.`codartic`)"
            C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
            If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafac1", "scaalb")
                Codigo = Replace(Codigo, "scafac", "scaalb")
                Codigo = Replace(Codigo, "slifac", "slialb")
                
                C = C & " WHERE " & Codigo
            End If
            C = C & ")"
    
    Else
    
        'Ventas
        C = "(SELECT Count(*) from (`scafpc` `scafpc` INNER JOIN `scafpa` `scafpa`"
        C = C & " ON ((`scafpc`.`codprove`=`scafpa`.`codprove`) AND (`scafpc`.`fecfactu`=`scafpa`.`fecfactu`))"
        C = C & " AND (`scafpc`.`numfactu`=`scafpa`.`numfactu`)) INNER JOIN ((`sartic` `sartic` INNER JOIN"
        C = C & " `slifpc` `slifpc` ON `sartic`.`codartic`=`slifpc`.`codartic`)"
        C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`)"
        C = C & " ON (((`scafpa`.`codprove`=`slifpc`.`codprove`) AND (`scafpa`.`numfactu`=`slifpc`.`numfactu`))"
        C = C & " AND (`scafpa`.`fecfactu`=`slifpc`.`fecfactu`)) AND (`scafpa`.`numalbar`=`slifpc`.`numalbar`)"
        If Codigo <> "" Then C = C & " WHERE " & Codigo
        C = C & ") + ("

        C = C & " SELECT count(*)"
        C = C & " FROM   ((`scaalp` `scaalp` INNER JOIN `slialp` `slialp` ON ((`scaalp`.`numalbar`=`slialp`.`numalbar`) AND (`scaalp`.`fechaalb`=`slialp`.`fechaalb`)) AND (`scaalp`.`codprove`=`slialp`.`codprove`))"
        C = C & " INNER JOIN `sartic` `sartic` ON `slialp`.`codartic`=`sartic`.`codartic`)"
        C = C & " INNER JOIN `sfamia` `sfamia` ON `sartic`.`codfamia`=`sfamia`.`codfamia`"
        If Codigo <> "" Then
                Codigo = Replace(Codigo, "scafpa", "scaalp")
                Codigo = Replace(Codigo, "scafpc", "scaalp")
                Codigo = Replace(Codigo, "slifac", "slialp")
                
                C = C & " WHERE " & Codigo
        End If
        C = C & ")"
    End If
    
    
    C = "Select " & C & " AS total"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If DBLet(miRsAux.Fields(0), "N") > 0 Then ContadorDelUnion = True
    End If
    miRsAux.Close
    Codigo = ""
End Function


Private Sub cmdAceptarEtiqProv_Click()
'305: Listado para etiquetas de proveedor
'306: Listado para cartas a proveedor
Dim campo As String

    InicializarVbles
    
    'si es listado de CARTAS/eMAIL a proveedores comprobar que se ha seleccionado
    'una carta para imprimir
    If OpcionListado = 306 Then
        If txtcodigo(63).Text = "" Then
            MsgBox "Debe seleccionar una carta para imprimir.", vbInformation
            Exit Sub
        End If
        
        'Parametro cod. carta
        cadParam = "|pCodCarta= " & txtcodigo(63).Text & "|"
        numParam = numParam + 1
        
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveCarta.rpt"
        Titulo = "Cartas a Proveedores"
        conSubRPT = True
        
    Else 'ETIQUETAS
        cadParam = "|"
    
        'Nombre fichero .rpt a Imprimir
        nomRPT = "rComProveEtiq.rpt"
        Titulo = "Etiquetas de Proveedores"
        conSubRPT = False
    End If
    
    '====================================================
    '================= FORMULA ==========================
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
     If txtcodigo(58).Text <> "" Or txtcodigo(59).Text <> "" Then
        campo = "{sprove.codprove}"
        'Parametro Desde/Hasta Proveedor
        If Not PonerDesdeHasta(campo, "N", 58, 59, "") Then Exit Sub
    End If
    
    'Cadena para seleccion D/H COD. POSTAL
    '--------------------------------------------
     If txtcodigo(60).Text <> "" Or txtcodigo(61).Text <> "" Then
        campo = "{sprove.codpobla}"
        'Parametro Desde/Hasta cod. Postal
        If Not PonerDesdeHasta(campo, "T", 60, 61, "") Then Exit Sub
    End If
    
    '====================================================
        
        
    'Parametro a la Atencion de
    cadParam = cadParam & "pAtencion=""Att. " & txtcodigo(62).Text & """|"
    numParam = numParam + 1
    
    'ver si hay registros seleccionados para mostrar en el informe
    If Not HayRegParaInforme("sprove", cadSelect) Then Exit Sub
    
    Set frmMen = New frmMensajes
    frmMen.cadWHERE = cadSelect
    frmMen.OpcionMensaje = 9 'Etiquetas proveedores
    frmMen.Show vbModal
    Set frmMen = Nothing
    If cadSelect = "" Then Exit Sub
    
    If OpcionListado = 306 And Me.chkEMAIL(0).Value = 1 Then
        'Enviarlo por e-mail
        EnviarEMailMulti cadSelect, Titulo, "rComProveCarta.rpt", "sprove" 'email para proveedores
    Else
        LlamarImprimir False, False
    End If
    
End Sub


Private Sub cmdAceptarFacRect_Click()
Dim Cad As String
Dim TipoM As String * 3


    'Comprobar que se introdujo el motivo por el que se rectifica la factura
    If Trim(txtcodigo(87).Text) = "" Then
        MsgBox "Debe introducir el motivo de rectificación.", vbExclamation
        PonerFoco txtcodigo(87)
        Exit Sub
    End If


    TipoM = Mid(Me.cboTipomov(0).List(Me.cboTipomov(0).ListIndex), 1, 3)
    
    '[Monica]06/05/2015: publicidad clientes
    If TipoM = "FPC" Then
        'comprobar que existe la factura en tabla "scafac"
        Cad = "select count(*) from scafaccli where codtipom='" & TipoM & "' AND numfactu="
        Cad = Cad & txtcodigo(71).Text & " AND fecfactu=" & DBSet(txtcodigo(72).Text, "F")
        If RegistrosAListar(Cad) = 0 Then
            Cad = vbCrLf & String(40, "*") & vbCrLf
            Cad = Cad & vbCrLf & "No existe la factura que quiere rectificar" & vbCrLf & "¿Continuar?" & Cad
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    Else
        'comprobar que existe la factura en tabla "scafac"
        Cad = "select count(*) from scafac where codtipom='" & TipoM & "' AND numfactu="
        Cad = Cad & txtcodigo(71).Text & " AND fecfactu=" & DBSet(txtcodigo(72).Text, "F")
        If RegistrosAListar(Cad) = 0 Then
            Cad = vbCrLf & String(40, "*") & vbCrLf
            Cad = Cad & vbCrLf & "No existe la factura que quiere rectificar" & vbCrLf & "¿Continuar?" & Cad
            If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
    End If
    
    'Llegado aqui pongo los datos
    'si existe devolver estos datos para recuperla en el formulario de Albaranes
    Cad = TipoM & "|"
    Cad = Cad & txtcodigo(71).Text & "|"
    Cad = Cad & txtcodigo(72).Text & "|"
    Cad = Cad & QuitarCaracterEnter(txtcodigo(87).Text) & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
    
End Sub


Private Sub cmdAceptarGenPed_Click()
'Solicitar datos para Generar Pedido a partir de una Oferta
Dim Cad As String

    Cad = txtcodigo(24).Text & "|"
    Cad = Cad & txtcodigo(25).Text & "|"
    Cad = Cad & txtcodigo(26).Text & "|"
    Cad = Cad & txtnombre(4).Text & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub cmdAceptarHco_Click()
'pedir datos para Pasar de Albaranes a historico
Dim Cad As String

    'comprobar que todos los camos tienen valor
    If txtcodigo(50).Text = "" Or txtcodigo(51).Text = "" Or txtcodigo(52).Text = "" Then
        MsgBox "Debe rellenar todos los campos para pasar al histórico.", vbInformation
        Exit Sub
    End If

    'datos a devolver
    Cad = txtcodigo(50).Text & "|"
    Cad = Cad & txtcodigo(51).Text & "|"
    Cad = Cad & txtcodigo(52).Text & "|"
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub PonerParamCadOferta2()
Dim C As String
Dim L As Boolean
    Set miRsAux = New ADODB.Recordset
    If Me.txtcodigo(3).Text = "" And txtcodigo(4).Text = "" Then
        
        L = False
    Else
        
        C = "numofert <> " & txtcodigo(1).Text
        C = C & " AND codclien = " & CodClien
        If txtcodigo(3).Text <> "" Then C = C & " AND fecofert >='" & Format(txtcodigo(3).Text, FormatoFecha) & "'"
        If txtcodigo(4).Text <> "" Then C = C & " AND fecofert <='" & Format(txtcodigo(4).Text, FormatoFecha) & "'"
        L = True
    End If
    
    CadenaDesdeOtroForm = ""
    If L Then
        C = "Select * from " & NomTabla & " where " & C
        miRsAux.Open C, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRsAux.EOF Then
            frmListado2.Opcion = 21
            frmListado2.Show vbModal
            
        End If
        miRsAux.Close
    End If
    
    CadenaDesdeOtroForm = "{" & NomTabla & ".numofert} IN [" & txtcodigo(1).Text & CadenaDesdeOtroForm & "]"
End Sub

Private Sub cmdAceptarPedCom_Click()
'55: Informe Pedido de Compras (a Proveedor)
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim devuelve As String, campo As String
Dim CodPed As String
Dim campo1 As String, campo2 As String, campo3 As String
    
    
    
    
    If txtcodigo(73).Text = "" Then 'Nº del Pedido
        MsgBox "Debe seleccionar un Pedido para Imprimir.", vbInformation
        PonerFoco txtcodigo(73)
        Exit Sub
    Else
        NumCod = txtcodigo(73).Text
    End If
    
    If (OpcionListado = 239) And txtcodigo(76).Text = "" Then
        MsgBox "Debe seleccionar un Pedido y Fecha para Imprimir.", vbInformation
        PonerFoco txtcodigo(76)
        Exit Sub
    End If
    
    
    InicializarVbles
    conSubRPT = True
    CadenaParaEnvioMail = ""
    '===================================================
    '============ PARAMETROS ===========================
    Select Case OpcionListado
        Case 38
            indRPT = 7 '7: Pedidos de Clientes
            Titulo = "Pedido de Ventas"
        Case 239
            indRPT = 8 '8: Pedidos de Clientes (Historico)
            Titulo = "Hist. Pedido de Venta"
        Case 55
            indRPT = 14 '14: Pedidos a Proveedores
            Titulo = "Pedidos de Compras"
        Case 56
            indRPT = 15
            Titulo = "Hist. Pedidos de Compras"
    End Select
    
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomRPT, vImprimedirecto, cadPDFrpt) Then Exit Sub
     
    If OpcionListado = 38 Or OpcionListado = 239 Then
        campo1 = "numpedcl"
        campo2 = "fecpedcl"
        campo3 = "codclien"
    Else
        campo1 = "numpedpr"
        campo2 = "fecpedpr"
        campo3 = "codprove"
    End If
    
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de PEDIDO
    '--------------------------------------------
    If NumCod <> "" Then
        devuelve = "{" & NomTabla & "." & campo1 & "}=" & Val(NumCod)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        cadSelect = cadFormula
        
        If OpcionListado = 239 Then 'historico ( hay fecha)
            devuelve = "{" & NomTabla & "." & campo2 & "}= Date(" & Year(txtcodigo(76).Text) & "," & Month(txtcodigo(76).Text) & "," & Day(txtcodigo(76).Text) & ")"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
            devuelve = NomTabla & "." & campo2 & "='" & Format(txtcodigo(76).Text, FormatoFecha) & "'"
            If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        End If
        
        'Seleccionar otros PEdidos entre esas FEchas
        If Not (txtcodigo(74).Text = "" And txtcodigo(75).Text = "") Then
            campo = "{" & NomTabla & "." & campo2 & "}"
            devuelve = CadenaDesdeHasta(txtcodigo(74).Text, txtcodigo(75).Text, campo, "F")
            If devuelve = "Error" Then Exit Sub
            If cadFormula <> "" Then
                cadFormula = "(" & cadFormula & " OR " & devuelve & ")"
                cadSelect = "((" & cadSelect & ") OR " & CadenaDesdeHastaBD(txtcodigo(74).Text, txtcodigo(75).Text, campo, "F") & ")"
            Else
                cadFormula = devuelve
                cadSelect = CadenaDesdeHastaBD(txtcodigo(74).Text, txtcodigo(75).Text, campo, "F")
            End If
        
            'Filtrar solo los Pedidos del CLIENTE/PROVEEDOR que las solicita
            If CodClien <> "" Then
                campo = "{" & NomTabla & "." & campo3 & "}=" & CodClien
                If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
                If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
            End If
     
        End If
        
        'FALTA## para hco pedidos
        If OpcionListado = 38 Or OpcionListado = 239 Then
            CadenaParaEnvioMail = "3|" & CodClien & "|" & txtcodigo(73).Text & "|"
        Else
            'Proveedores
            CadenaParaEnvioMail = "51|" & CodClien & "|" & txtcodigo(73).Text & "|"
        End If
        
        
        
    Else
'        'Comprobar si se imprimen varios Pedidos
'        If txtCodigo(3).Text <> "" Or txtCodigo(4).Text <> "" Then
'         'Cadena para seleccion Desde y Hasta FECHA
'         '--------------------------------------------
'            campo = "{" & NomTabla & ".fecpedcl}"
'            devuelve = CadenaDesdeHasta(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'            If devuelve = "Error" Then Exit Sub
'            If Not AnyadirAFormula(cadFormula, devuelve) Then
'                Exit Sub
'            Else
'                devuelve = CadenaDesdeHastaBD(txtCodigo(3).Text, txtCodigo(4).Text, campo, "F")
'                If devuelve = "Error" Then Exit Sub
'                If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
'            End If
'        End If
    End If
    
    If OpcionListado = 38 Or OpcionListado = 239 Then
        'Añadimos a los parametros el tipo de IVA que se aplica a ese cliente (para saber si esta exento o no de IVA)
        devuelve = DevuelveDesdeBDNew(conAri, "sclien", "tipoiva", "codclien", CodClien, "N")
        If devuelve <> "" Then
            cadParam = cadParam & "pTipoIVA=" & devuelve & "|"
            numParam = numParam + 1
        End If
        
        'PORTES
        cadParam = cadParam & "vPortes=""" & vParamAplic.ArtPortes & """|"
        numParam = numParam + 1
    End If

    'comprobar que hay datos para mostrar en el Informe
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
    
    LlamarImprimir True, False, CadenaParaEnvioMail
End Sub


' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
Private Sub cmdAceptarPedConfirma_Click()
'Confirmacion entrega del pedido
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim NomPDF As String 'Nombre del fichero .pdf
Dim campo As String
Dim RS As ADODB.Recordset

    If txtcodigo(116).Text = "" Then
        MsgBox "Debe seleccionar una carta para Imprimir la Confirmación entrega del Pedido.", vbInformation
        PonerFoco txtcodigo(116)
        Exit Sub
    End If
    
    
    PrepararCarpetasEnvioMail True
    
    InicializarVbles
    
    'Se pasa como parametro la carta a imprimir
    If Me.txtcodigo(116).Text <> "" Then
        cadParam = cadParam & "|pCodCarta=" & CInt(Me.txtcodigo(116).Text) & "|"
    Else
        cadParam = cadParam & "|pCodCarta=" & CInt(0) & "|"
    End If
    numParam = numParam + 1
    
    
    indRPT = 40 'Añade los parametros de la tabla scrystal para el informe
    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu, vImprimedirecto, pPdfRpt) Then
        Exit Sub
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    NomTabla = "scaped"
    
    'Cadena para seleccion Clientes de Pedido
    '--------------------------------------------
    If txtcodigo(114).Text <> "" Then
        campo = "{scaped.numpedcl}=" & txtcodigo(114).Text
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        cadSelect = cadFormula
    End If
       
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Not HayRegParaInforme(NomTabla, cadSelect) Then Exit Sub
       
       
    If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
       
'    LlamarImprimir

     With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = True
        .Opcion = 238
        .Titulo = "Confirmación entrega Pedido"
        .NombreRPT = nomDocu
        .NombrePDF = pPdfRpt
        .ConSubInforme = True
        .Show vbModal
    End With
    
    
    'FALTA###
'    Exit Sub
    If Dir(App.Path & "\docum.pdf", vbArchive) = "" Then
        MsgBox "No se encuentra el archivo", vbExclamation
        Exit Sub
    End If
    NomPDF = App.Path & "\Temp\PEV-" & Format(NumCod, "0000000") & ".pdf"
    FileCopy App.Path & "\docum.pdf", NomPDF
    
    'Obtener los ficheros que hay en el directorio de documentos
'    MiRuta = "" & App.Path & "" & "\PDF-Docum\"


    '-- obtener los datos para envio e-mail
    campo = "SELECT numpedcl,fecpedcl,codclien,nomclien,mailconfir"
    campo = campo & " FROM " & NomTabla & " WHERE numpedcl=" & NumCod
    Set RS = New ADODB.Recordset
    RS.Open campo, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
    campo = ""
    If Not RS.EOF Then
        If DBLet(RS!mailconfir, "T") <> "" Then campo = RS!nomclien & "|" & RS!mailconfir & "|"
    End If
    
    RS.Close
    Set RS = Nothing

    If campo = "" Then MsgBox "No hay dirección e-mail en el pedido para enviar confirmación de entrega.", vbExclamation
    
    If Dir(NomPDF, vbArchive) <> "" And campo <> "" Then
    
        '- añadir el subject del e-mail
        campo = campo & "Confirmación entrega pedido " & vEmpresa.nomempre & "|"
        '- añadir el cuerpo del mensaje
        campo = campo & "Le confirmamos que su pedido adjunto Nº " & NumCod & " de fecha " & FecEntre & " le será entregado en la semana "
        campo = campo & DevuelveDesdeBDNew(conAri, NomTabla, "sementre", "numpedcl", NumCod, "N") & ".|"
        
        'El adjunto, para que no se llame docum.pdf
        campo = campo & NomPDF & "|"
        
        frmEMail.DatosEnvio = campo
        frmEMail.Opcion = 0 'Envio documento
        frmEMail.Show vbModal
    
        If frmEMail.DatosEnvio = "OK" Then
            campo = "UPDATE " & NomTabla & " SET envconfir=1"
            campo = campo & " WHERE numpedcl=" & NumCod
            conn.Execute campo
        End If
        frmEMail.DatosEnvio = ""
        
    End If
    
    'If Dir(NomPDF, vbArchive) <> "" Then Kill NomPDF
End Sub
' ----



Private Sub cmdAceptarPte_Click()
'LIstado Material Pendiente de recibir
Dim Codigo As String
Dim Cad As String

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    'Pasar el ORDEN del informe como parametro
    If OpcionListado = 307 Then
        If Me.OptOrdenArt Then
            Cad = "{slippr.codartic}"
        Else
            Cad = "{scappr.numpedpr}"
        End If
        cadParam = cadParam & "pOrden=" & Cad & "|"
        numParam = numParam + 1
    End If
    
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe
    
    'Cadena para seleccion D/H PROVEEDOR
    '--------------------------------------------
    If txtcodigo(65).Text <> "" Or txtcodigo(66).Text <> "" Then
        Codigo = "{scappr.codprove}"
        If OpcionListado = 308 Then Codigo = "{scaalp.codprove}"
        Cad = "pDHProveedor=""Proveedor: "
        If Not PonerDesdeHasta(Codigo, "N", 65, 66, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(69).Text <> "" Or txtcodigo(70).Text <> "" Then
        Codigo = "{scappr.fecpedpr}"
        If OpcionListado = 308 Then Codigo = "{scaalp.fechaalb}"
        Cad = "pDHFecha=""Fecha Ped.: "
        If OpcionListado = 308 Then Cad = "pDHFecha=""Fecha Alb.: "
        If Not PonerDesdeHasta(Codigo, "F", 69, 70, Cad) Then Exit Sub
    End If
    
    If OpcionListado = 307 Then '307: List. Materia pendiente de recibir
        'Cadena para seleccion D/H ARTICULO
        '--------------------------------------------
        If txtcodigo(67).Text <> "" Or txtcodigo(68).Text <> "" Then
            Codigo = "{slippr.codartic}"
            Cad = "pDHArticulo=""Artículo: "
            If Not PonerDesdeHasta(Codigo, "T", 67, 68, Cad) Then Exit Sub
        End If
    End If
    
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If OpcionListado = 307 Then
        Cad = "scappr INNER JOIN slippr ON scappr.numpedpr=slippr.numpedpr "
        Titulo = "Material Pendiente de recibir"
        nomRPT = "rComPteRecibir.rpt"
    Else
        Cad = "scaalp INNER JOIN slialp ON scaalp.numalbar=slialp.numalbar AND scaalp.fechaalb=slialp.fechaalb AND scaalp.codprove=slialp.codprove "
        Titulo = "Pendiente de Factura"
        nomRPT = "rComPteFactura.rpt"
    End If
    
    If Not HayRegParaInforme(Cad, cadSelect) Then Exit Sub

    'Mostrar el Informe
    conSubRPT = False
    LlamarImprimir False, False
End Sub


Private Sub cmdAceptarReimpFac_Click()
'Reimprimir Facturas ya contabilizadas
Dim TipoM As String * 3
'Dim TipoMh As String * 3
Dim Codigo As String
Dim b As Boolean
Dim TipoFactura As Byte

    InicializarVbles
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe
    
    'Desde/Hasta tipo movimiento
    '---------------------------------------------
    TipoM = Mid(Me.cboTipomov(1).List(Me.cboTipomov(1).ListIndex), 1, 3)
    If TipoM <> "" Then
        Codigo = "({scafac.codtipom}='" & TipoM & "') "
'        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
        cadSelect = cadFormula
'        If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    End If

    
    'Desde/Hasta numero de FACTURA
    '---------------------------------------------
    If txtcodigo(83).Text <> "" Or txtcodigo(84).Text <> "" Then
        Codigo = "{scafac.numfactu}"
        If Not PonerDesdeHasta(Codigo, "N", 83, 84, "") Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(85).Text <> "" Or txtcodigo(86).Text <> "" Then
        Codigo = "{scafac.fecfactu}"
        If Not PonerDesdeHasta(Codigo, "F", 85, 86, "") Then Exit Sub
    End If
    
    If CBool(Me.chk_duplicado.Value) Then
        cadParam = "pDuplicado=1|"
    Else
        cadParam = "pDuplicado=0|"
    End If
    
    
    TipoFactura = 0
    Codigo = Mid(cboTipomov(1).Text, 1, 3)
    If Codigo <> "" Then
        If Codigo = "FTI" Then
            TipoFactura = 1                        'Facturas ticket
        Else
            If Codigo = "FAZ" Then TipoFactura = 2 'FAacturas B
        End If
    End If
    
    
    ImprimirFacturas cadFormula, cadParam, cadSelect, TipoFactura
    
End Sub

Private Sub cmdAceptarTrasHco_Click()
Dim devuelve As String
Dim Cad As String
'IMPRIME INFORME y DESPUES PREGUNTA SI TRASPASAR AL HISTORICO

    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '===================================================
    '================= FORMULA =========================
    'será la cadena WHERE para el Informe y la SQL del Traspaso a Hco
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtcodigo(43).Text <> "" Or txtcodigo(44).Text <> "" Then
        Codigo = "{scapre.codclien}"
        Cad = "pDHCliente=""Cliente: "
        If Not PonerDesdeHasta(Codigo, "N", 43, 44, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion AGENTE
    '--------------------------------------------
    If txtcodigo(45).Text <> "" Or txtcodigo(46).Text <> "" Then
        Codigo = "{scapre.codagent}"
        Cad = "pDHAgente=""Agente: "
        If Not PonerDesdeHasta(Codigo, "N", 45, 46, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta FECHA
    '--------------------------------------------
    If txtcodigo(22).Text <> "" Or txtcodigo(23).Text <> "" Then
        Codigo = "{scapre.fecofert}"
        Cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(Codigo, "F", 22, 23, Cad) Then Exit Sub
    End If
    
    'Cadena para seleccion Desde y Hasta Nº OFERTA
    '---------------------------------------------
    If txtcodigo(20).Text <> "" Or txtcodigo(21).Text <> "" Then
        Codigo = "{scapre.numofert}"
        Cad = "pDHOferta=""Nº Oferta: "
        If Not PonerDesdeHasta(Codigo, "N", 20, 21, Cad) Then Exit Sub
    End If
    
    'Seleccionar para estos criterios solo las Ofertas que no esten Aceptadas
    '------------------------------------------------------------------------
    devuelve = " {scapre.aceptado} = 0 "
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub 'Para Crystal
    If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub 'Para MySQL
    
    'Comprobar que hay datos que mostrar antes de Abrir el Informe
    If Not HayRegParaInforme("scapre", cadSelect) Then Exit Sub

    'Mostrar el Informe
    LlamarImprimir False, False
    
    'Preguntar si Traspasamos los Datos seleccionados al Histórico
    If MsgBox("¿Desea pasar estas Ofertas al Histórico?", vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        If TraspasoOfertaAHco(cadSelect) Then MsgBox "Traspaso de Ofertas a Histórico realizado correctamente. ", vbInformation
    End If
End Sub

Private Sub cmdBajar_Click()
    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


'Private Sub cmdEnvioMail_Click()
'Dim RS As ADODB.Recordset
'
'
'    'El proceso constara de varias fases.
'    'Fase 1: Montar el select y ver si hay registros
'    'Fase 2: Preparar carpetas para los pdf
'    'Fase 3: Generar para cada factura (una a una) del select su pdf
'    'Fase 4: Enviar por mail, adjuntando los archivos correspondientes
'
'    If Text1(0).Text = "" Then
'        MsgBox "Ponga el asunto", vbExclamation
'        Exit Sub
'    End If
'
'    'Cadena para seleccion Desde y Hasta FECHA
'    '--------------------------------------------
'
'    'AHora pongo los tipo de facturas
'    cadFormula = ""
'    cadSelect = ""  'ME dira si estan todas o no
'    For indCodigo = 0 To Me.ListTipoMov(1000).ListCount - 1
'        If Me.ListTipoMov(1000).Selected(indCodigo) Then
'            'Esta checkeado
'            cadFormula = cadFormula & " OR scafac.codtipom = '" & Trim(Mid(ListTipoMov(1000).List(indCodigo), 1, 3)) & "'"
'        Else
'            cadSelect = "NO"
'        End If
'    Next indCodigo
'
'    If cadFormula = "" Then
'        MsgBox "Seleccione algun tipo de factura", vbExclamation
'        Exit Sub
'    Else
'        cadFormula = Mid(cadFormula, 4)
'    End If
'    If cadSelect = "" Then
'        'Significa que estan todos. No tiene sentido poner que codtipo='fr or codtipo='FT  ESTAN TODAS
'        cadFormula = " scafac.codtipom <> 'FTI'"
'    End If
'    'En notabla tendre
'
'    NomTabla = "(" & cadFormula & ")"
'
'    InicializarVbles
'    cadFormula = ""
'    cadSelect = ""
'    If txtCodigo(110).Text <> "" Or txtCodigo(111).Text <> "" Then
'        Codigo = "scafac.codclien"
'        If Not PonerDesdeHasta(Codigo, "N", 110, 111, "") Then Exit Sub
'    End If
'
'    If txtCodigo(108).Text <> "" Or txtCodigo(109).Text <> "" Then
'        Codigo = "scafac.fecfactu"
'        If Not PonerDesdeHasta(Codigo, "F", 108, 109, "") Then Exit Sub
'    End If
'
'    If txtCodigo(106).Text <> "" Or txtCodigo(107).Text <> "" Then
'        Codigo = "scafac.numfactu"
'        If Not PonerDesdeHasta(Codigo, "N", 106, 107, "") Then Exit Sub
'    End If
'
'
'    Screen.MousePointer = vbHourglass
'
'    'Eliminamos temporales
'    conn.Execute "DELETE from tmpnlotes where codusu =" & vUsu.Codigo
'
'    If cadSelect <> "" Then cadSelect = cadSelect & " AND "
'    cadSelect = cadSelect & NomTabla
'    cadSelect = " WHERE " & cadSelect
'
'    Set RS = New ADODB.Recordset
'    DoEvents
'
'
'
'    'Ahora insertare en la tabla temporal tminformes las facturas que voy a generar pdf
'    Codigo = "insert into tmpnlotes (codusu,numalbar,codprove,codartic,numlinea,fechaalb,codalmac,cantidad) "
'    Codigo = Codigo & " values ( " & vUsu.Codigo & ",'"
'
'    If Not PrepararCarpetasEnvioMail Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'
'
'    'Vamos a meter todas las facturas en la tabla temporal para comprobar si tienen mail
'    'los clientes
'
'    NomTabla = "Select codtipom,numfactu,codclien,fecfactu,totalfac from scafac  " & cadSelect
'    'El orden vamos a hacerlo por: Tipo documento
'    NomTabla = NomTabla & " ORDER BY codtipom"
'    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    NumRegElim = 0
'    While Not RS.EOF
'        NomTabla = RS!codtipom & "'," & RS!codClien & "," & RS!NumFactu & "," & CStr(RS!NumFactu Mod 32000) & ",'" & Format(RS!FecFactu, FormatoFecha)
'
'        'El tipo de informe lo guardare en el ultimo campo
'        'El report es el = 12
'        NomTabla = NomTabla & "',12," & TransformaComasPuntos(CStr(DBLet(RS!TotalFac, "N"))) & ")"
'        conn.Execute Codigo & NomTabla
'        NumRegElim = NumRegElim + 1
'        RS.MoveNext
'    Wend
'    RS.Close
'
'
'    If NumRegElim = 0 Then
'        MsgBox "Ningun dato a enviar por mail", vbExclamation
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If
'
'    '--------------------------------------------------------------------------------------------------
'    '
'    'Ahora cojemos las facturas que son FVA pero tienen numero terminal. COn el desde /hasta seleccionado
'    'MIRAMOS en la tabla scafac1
'
'    'Compruebo si tiene codclien
'    NomTabla = "select scafac1.* from scafac1 ,scafac where scafac1.codtipom=scafac.codtipom and scafac1.numfactu=scafac.numfactu and scafac1.fecfactu =scafac.fecfactu"
'    'NomTabla = "Select codtipom,numfactu,fecfactu from scafac1   " & cadSelect
'    'El cad select LLEVA el where.  Se lo quito
'    cadSelect = Mid(cadSelect, 7)
'    NomTabla = NomTabla & " AND " & cadSelect & "  AND numtermi>=0  "
'
'    RS.Open NomTabla, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not RS.EOF
'        NomTabla = "numalbar = '" & RS!codtipom & "' AND fechaalb = '" & Format(RS!FecFactu, FormatoFecha) & "' AND numlinea = " & CStr(RS!NumFactu Mod 32000)
'        'El tipo de informe lo guardare en el ultimo campo
'        'El report es el = 12
'        NomTabla = "UPDATE tmpnlotes SET codalmac = 18 WHERE codusu = " & vUsu.Codigo & " AND " & NomTabla
'        conn.Execute NomTabla
'
'
'        RS.MoveNext
'    Wend
'    RS.Close
'    'Numero de registros
'    NomTabla = NumRegElim
'
'    'AHora ya tengo todos los datos de las facturas que voy  a imprimir
'    'Entonces copruebo si para los clientes si tienen puesto el campo mail o no
'    If optEnvioMail(0).Value Then
'        'Selecciona mail comercial
'        cadSelect = "2"  'de maiclie2
'    Else
'        cadSelect = "1"  'de maiclie1
'    End If
'    cadSelect = "Select codclien,maiclie" & cadSelect
'    cadSelect = cadSelect & " as email from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
'    cadSelect = cadSelect & " group by codclien having email is null"
'    RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    NumRegElim = 0
'    While Not RS.EOF
'        NumRegElim = NumRegElim + 1
'        RS.MoveNext
'    Wend
'    RS.Close
'
'    If NumRegElim > 0 Then
'        If MsgBox("Tiene cliente sin mail. Continuar sin sus datos?", vbQuestion + vbYesNo) = vbNo Then
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'
'        'Si no salimos borramos
'        RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cadSelect = "DELETE from tmpnlotes where codusu =" & vUsu.Codigo & " and codprove ="
'        While Not RS.EOF
'            conn.Execute cadSelect & RS!codClien
'            RS.MoveNext
'        Wend
'        RS.Close
'
'
'        cadSelect = "Select count(*) from tmpnlotes where codusu =" & vUsu.Codigo
'        RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        NumRegElim = 0
'        If Not RS.EOF Then
'            If Not IsNull(RS.Fields(0)) Then NumRegElim = DBLet(RS.Fields(0), "N")
'
'        End If
'        RS.Close
'
'        If NumRegElim = 0 Then
'            'NO hay datos para enviar
'
'            Screen.MousePointer = vbDefault
'            MsgBox "No hay datos para enviar por mail", vbExclamation
'            Exit Sub
'        Else
'            cadSelect = "Hay " & NumRegElim & " facturas para enviar por mail." & vbCrLf & "¿Continuar?"
'            If MsgBox(cadSelect, vbQuestion + vbYesNo) = vbNo Then NumRegElim = 0
'        End If
'        If NumRegElim = 0 Then
'            Set RS = Nothing
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
'        NomTabla = NumRegElim
'
'
'
'
'    End If
'
'    PonerTamnyosMail True
'    frmPpal.visible = False
'    'Voy arriesgar.
'    'Confio en que no envien por mail mas de 32000 facturas (un integer)
'    Label4(22).Caption = "Preparando datos"
'    Me.ProgressBar1.Max = CInt(NomTabla)
'    Me.ProgressBar1.Value = 0
'
'
'
'    NumRegElim = 0
'    If GeneracionEnvioMail(RS) Then NumRegElim = 1
'
'
'    'Si ha ido todo bien entonces numregelim=1
'    If NumRegElim = 1 Then
'        'Procederemos a enviarlos por mail
'        If optEnvioMail(0).Value Then
'            'Selecciona mail comercial
'            cadSelect = "2"  'de maiclie2
'        Else
'            cadSelect = "1"  'de maiclie1
'        End If
'        cadSelect = "Select nomclien,maiclie" & cadSelect
'        cadSelect = cadSelect & " as email,tmpnlotes.* from tmpnlotes,sclien where codusu = " & vUsu.Codigo & " and codclien=codprove"
''        cadSelect = cadSelect & " group by codclien having email is null"
'
'
'        frmEMail.DatosEnvio = Text1(0).Text & "|" & Text1(1).Text & "|" & Abs(chkMail.Value) & "|" & cadSelect & "|"
'        frmEMail.Opcion = 4 'Multienvio de facturacion
'        frmEMail.Show vbModal
'
'
'        'Para tranquilizar las pantallas, borrar los ficheros generados
'        'Confio en que no envien por mail mas de 32000 facturas (un integer)
'        Label14(22).Caption = "Restaurando ...."
'        Me.ProgressBar1.visible = False
'        Me.Refresh
'        DoEvents
'        Espera 1
'        PrepararCarpetasEnvioMail
'        Me.ProgressBar1.visible = True
'
'
'    End If
'
'
'
'
'    'Es para evitar la cantidad de pantallas abriendose y cerrandose
'    Me.visible = False
'    PonerTamnyosMail False
'    Espera 1
'    Unload Me
'    frmPpal.Show
'
'    Screen.MousePointer = vbDefault
'End Sub
        
        
        
Private Function GeneracionEnvioMail(ByRef RS As ADODB.Recordset) As Boolean

    On Error GoTo EGeneracionEnvioMail
    GeneracionEnvioMail = False

    
    cadSelect = "Select * from tmpnlotes where codusu =" & vUsu.Codigo & " ORDER BY codalmac,numalbar,codprove"
    RS.Open cadSelect, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    CodClien = ""
    While Not RS.EOF
        
        If Dir(App.Path & "\docum.pdf", vbArchive) <> "" Then Kill App.Path & "\docum.pdf"
    
        Label14(22).Caption = "Factura: " & RS!NumAlbar & " " & RS!codArtic
        Label14(22).Refresh
        
        If CodClien <> RS!codAlmac Then   'If CodClien <> RS!codTipoM Then
            'OTRO TIPO DE DOCUMENTO
            
            '''''If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then
            If Not PonerParamRPT(RS!codAlmac, cadParam, numParam, NumCod, vImprimedirecto, cadPDFrpt) Then
                Exit Function
            End If
            CodClien = RS!codAlmac
        End If
        cadFormula = "({scafac.codtipom}='" & RS!NumAlbar & "') "
        cadFormula = cadFormula & " AND ({scafac.numfactu}=" & RS!codArtic & ") "
        cadFormula = cadFormula & " AND ({scafac.fecfactu}= Date(" & Year(RS!FechaAlb) & "," & Month(RS!FechaAlb) & "," & Day(RS!FechaAlb) & "))"


          
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = cadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = True
            .NombreRPT = NumCod
            .NombrePDF = cadPDFrpt
            .Opcion = 53
            .Titulo = ""
            .Show vbModal
        End With
    
                    
        'Subo el progress bar
        Label14(22).Caption = "Generando PDF"
        Label14(22).Refresh
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        If (Me.ProgressBar1.Value Mod 25) = 24 Then
            Me.Refresh
            DoEvents
            Espera 1
        End If
        Me.Refresh
        DoEvents
        
        
        
        'FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codProve, "0000000") & Format(RS!codArtic, "0000000") & Format(RS!FechaAlb, "yymmdd") & ".pdf"
        FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & RS!NumAlbar & Format(RS!codArtic, "0000000") & ".pdf"
        
        RS.MoveNext
    Wend
    RS.Close
    
    Set RS = Nothing
    GeneracionEnvioMail = True
    Exit Function
EGeneracionEnvioMail:
       MuestraError Err.Number
End Function

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub









Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 31, 35 '31: Informe Ofertas
                        '35: Informe Historico Ofertas
                PonerFoco txtcodigo(1)
            Case 32, 33 '32: Recordatorio de Oferta
                        '33: Informe Valoracion de Oferta
                PonerFoco txtcodigo(5)
            Case 34, 92 '34: Informe Ofertas Efectuadas
                        '92: Informe Gastos técnicos
                PonerFoco txtcodigo(16)
            Case 36 '36: Traspaso Ofertas a Historico
                PonerFoco txtcodigo(43)
            Case 37 '37: Generar Pedido de OFerta
                PonerFoco txtcodigo(24)
            Case 40 '40: Carta Confirmacion de Pedido
                PonerFoco txtcodigo(77)
            Case 46, 48, 90, 91 '46: Informe Clientes Inactivos
                        '48: Informe de Altas de Nuevos Clientes
                        '90: Etiquetas de clientes
                        '91: Cartas a clientes
                PonerFoco txtcodigo(27)
            Case 190, 191
                        '190: Etiquetas de socios
                        '191: Cartas a socios
                PonerFoco txtcodigo(14)
            
            Case 47 '47: Informe de Clientes
                PonerFoco txtcodigo(33)
            Case 38, 239, 55, 56 '55: Informe de Pedido de Compras (proveedor)
                PonerFoco txtcodigo(73)
            Case 57 '57: Pasar Pedido a Albaran de Compras(Proveedores)
                PonerFoco txtcodigo(47)
            Case 80, 81 '80: Pasar albaranes al historico (ventas clientes)
                            '81: Pasar pedidos al historico (ventas clientes)
                PonerFoco txtcodigo(50)
            
            Case 225 'Datos para Factura Rectificativa
                PonerFoco txtcodigo(71)
            Case 226 'Datos para Reimprimir Facturas
                PonerFocoCbo Me.cboTipomov(1)
                
            Case 230 'Listado Ventas por Familia
                PonerFoco txtcodigo(96)
            
            Case 232 'Listado Facturacion por cliente
                PonerFoco txtcodigo(3)
                
            ' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
            Case 238 'Confirmacion entrega Pedido
                PonerFoco txtcodigo(116)
            ' ----
                
            Case 240 'Inf. Cierre caja TPV
                PonerFoco txtcodigo(88)
                
            Case 305, 306 '305: Listado Etiquetas proveedor
                          '306: Listado Cartas a proveedores
                PonerFoco txtcodigo(58)
            Case 307, 308 '307: List. Pendiente de Recibir (COMPRAS)
                          '308: List. Pendiente de Facturar (COMPRAS)
                PonerFoco txtcodigo(65)
                
            Case 310, 311, 312 'Listado Compras por Proveedor/Familia/Articulo
                                '312: Listado albaranes por proveedor
                PonerFoco txtcodigo(90)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single
Dim devuelve As String
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    indCodigo = 0
    NomTabla = ""
    
    'imgbuscar
    Me.imgBuscarOfer(0).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Me.imgBuscarOfer(4).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    For kCampo = 8 To 23
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 27 To 44
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 48 To 55
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 58 To 60
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    'fecha
    For kCampo = 0 To 1
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 11 To 19
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 23 To 32
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    
    'Ocultar todos los Frames de Formulario
    Me.FrameClienInactivos.visible = False
    Me.FrameClientes.visible = False
    Me.FrameGenAlbCom.visible = False
    Me.FramePasarHco.visible = False
    Me.FrameEtiqProv.visible = False
    Me.FramePteRecibir.visible = False
    Me.FrameFacRectif.visible = False
    Me.FrameFacReimprimir.visible = False
    Me.FramePedidos.visible = False
    ' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
    Me.FramePedConfirma.visible = False
    ' ----
    Me.FrameCierreCaja.visible = False
    Me.FrameCompras.visible = False
    Me.FrameEstVentasFam.visible = False
    Me.FrameEstCliente.visible = False
    Me.FrameSocios.visible = False
    
'    FrameEnvioFacMail.visible = False
    CommitConexion
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 31, 35 '31: Informe de Ofertas
                    '35: Informe Historico de Ofertas
            
        Case 32, 33 '32: Recordatorio de Ofertas
                    '33:Informe Valoración de Ofertas
            
        Case 34, 92 '34: Informe Ofertas Efectuadas
                    '92: Informe Gastos Técnicos
            
        Case 36 '36: Traspaso a Historico (IMPRIME LISTADO Y PREGUNTA SI TRASPASO A HCO)
            
        Case 37 '37: Pedir datos para pasar Oferta a Pedido (NO IMPRIME LISTADO)
        
        Case 40 '40: Cartas Confirmacion de Pedidos
        
        Case 46, 48, 90, 91 '46: Informe Clientes Inactivos
                        '90: Etiquetas de clientes
                        '91: Cartas a clientes
            PonerFrameClienInacVisible True, H, W
            indFrame = 5
            If OpcionListado = 90 Then
                CargarComboTipoMov 2
                FrameImpClien.visible = False
            End If
            
        Case 190, 191   '190: Etiquetas de socios
                        '191: Cartas a socios
            W = 7185
            H = 4425
            PonerFrameSociosVisible True, H, W
            indFrame = 5
            Me.Frame14.visible = (OpcionListado = 191)
            If (OpcionListado = 191) Then Me.Label1.Caption = "Cartas a Socios"
            
            Me.chkMembrete.Value = 0
            Me.chkMembrete.visible = (OpcionListado = 191)
            Me.chkMembrete.Enabled = (OpcionListado = 191)
            
        
        Case 47 '47: Informe de Clientes
            W = 8960
            H = 6020
            PonerFrameVisible Me.FrameClientes, True, H, W
            CargarListViewOrden
            indFrame = 6
            
        Case 38, 239, 55, 56
                '38: Pedidos Venta
                '55: Informe de Pedido de Compras (Proveedor)
                '56: Informe de Hist. Pedido de Compras (Proveedor)
            PonerFramePedVisible H, W
            indFrame = 12
            If NumCod <> "" Then txtcodigo(73).Text = NumCod
            
            
            
        Case 57 '57: Pedir datos para pasar de Pedido a Albaran (NO IMPRIME LISTADO)
            W = 6315
            H = 4455
            PonerFrameVisible Me.FrameGenAlbCom, True, H, W
            indFrame = 7
            Me.Caption = "Generar Albaran Compras"
            'Poner el trabajador conectado
            Me.txtcodigo(47).Text = PonerTrabajadorConectado(devuelve)
            Me.txtnombre(47).Text = devuelve
            Me.txtcodigo(49).Text = Format(Now, "dd/mm/yyyy")
        
        Case 80, 81 '80: pasar albaranes al historico (ventas)
                        '81: pasar pedidos al historico (ventas)
            H = 4575
            W = 6920
            PonerFrameVisible Me.FramePasarHco, True, H, W
            indFrame = 8
            Me.Caption = "Eliminar"
            Select Case OpcionListado
                Case 80, 82: Me.Label3(4).Caption = "Pasar Albaran al histórico"
                Case 81: Me.Label3(4).Caption = "Pasar Pedido al histórico"
            End Select
            Me.txtcodigo(50).Text = Format(Now, "dd/mm/yyyy")
            Me.txtcodigo(51).Text = PonerTrabajadorConectado(devuelve)
            Me.txtnombre(51).Text = devuelve
            
        Case 225 'Factura rectificativa
            H = 4420
            W = 5740
            PonerFrameVisible Me.FrameFacRectif, True, H, W
            indFrame = 11
            Me.Caption = "Facturas rectificativas"
            CargarComboTipoMov (0)
'            Me.cboTipomov(0).ListIndex = 2
            
        Case 226 'Reimprimir Factura
            H = 4455
            W = 6555
            PonerFrameVisible Me.FrameFacReimprimir, True, H, W
            indFrame = 14
            CargarComboTipoMov (1)
            
            
            cadFormula = DevuelveDesdeBDNew(conAri, "scryst", "nomcryst", "codcryst", "18", "N")
            Me.chkFormatoTPV.Value = 0
            If cadFormula = "" Then
                'NO SE HA ENCONTRADOR
                Me.chkFormatoTPV.Enabled = False
                cadFormula = "Formato NO encontrado"
            End If
            Me.chkFormatoTPV.Caption = cadFormula
            
'            CargarComboTipoMov (2)
            
        Case 230, 231 '230: Estadistica ventas por familia
                      '231: Detalle facturacion socios
            indFrame = 17
            H = 5805
            If OpcionListado = 231 Then
                H = 4325
                Me.cmdAceptarEstVentas.Top = 3400
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarEstVentas.Top
                Me.Label9(31).Caption = "Detalle Facturación Socios"
                CargarCombo
                Me.Combo1.ListIndex = 0
            End If
            W = 7035
            Me.Frame12.visible = (OpcionListado = 230)
            Me.Frame11.visible = (OpcionListado = 231)
            PonerFrameVisible Me.FrameEstVentasFam, True, H, W
            
        Case 232
           '232: Detalle facturacion clientes
            indFrame = 17
            H = 4325
            Me.cmdAceptarEstVentas.Top = 3400
            Me.cmdCancel(indFrame).Top = Me.cmdAceptarEstVentas.Top
            CargarCombo2
            Me.Combo2.ListIndex = 0
            W = 7035
            PonerFrameVisible Me.FrameEstCliente, True, H, W
             
            
        ' ---- [04/11/2009] [LAURA] : Añadir botón para enviar informe confirmacion entrega del Pedido
        Case 238 'Confirmacion entrega pedido
            W = 6315
            H = 4095
            PonerFrameVisible Me.FramePedConfirma, True, H, W
            indFrame = 19
            Me.Caption = "Confirmación entrega Pedido"
            If NumCod <> "" Then txtcodigo(114).Text = NumCod
            txtcodigo(115).Text = Format(FecEntre, "dd/mm/yyyy")
            BloquearTxt txtcodigo(114), True
            BloquearTxt txtcodigo(115), True
            
'            NomTabla = "scaped"
'            NomTablaLin = "sliped"
        ' ----
        
        Case 240 'Inf. cierre caja TPV
            H = 3800
            W = 6300
            PonerFrameVisible Me.FrameCierreCaja, True, H, W
            indFrame = 15
'            CargarComboTipoPago
'            Combo1.ListIndex = 0
            'Mostrar la fecha de hoy
            txtcodigo(88).Text = Format(Now, "dd/mm/yyyy")
            txtcodigo(89).Text = Format(Now, "dd/mm/yyyy")
            
        
        Case 305, 306 '305: Etiquetas de proveedor
                      '306: Cartas a proveedor
            indFrame = 9
            H = 5325
            W = 7835
            PonerFrameVisible Me.FrameEtiqProv, True, H, W
            Me.Frame2.visible = (OpcionListado = 306)
            If (OpcionListado = 306) Then Me.Label9(1).Caption = "Cartas a Proveedores"
            
        Case 307, 308 '307: List. Material Pendiente de recibir (COMPRAS)
                      '308: List. Albaranes ptes de facturar (COMPRAS)
            indFrame = 10
            If OpcionListado = 307 Then
                Me.Label9(19).Caption = "Material pendiente de recibir"
                H = 5200
            Else
                Me.Label9(19).Caption = "Albaranes pendiente de factura"
                H = 4200
                Me.cmdAceptarPte.Top = 3500
                Me.cmdCancel(10).Top = Me.cmdAceptarPte.Top
            End If
            W = 7035
            PonerFrameVisible Me.FramePteRecibir, True, H, W
            Me.Frame6.visible = (OpcionListado = 307)
            Me.Frame7.visible = (OpcionListado = 307)
            
        Case 310, 311, 312 '310: Listado COMPRAS por proveedor
                            '312: Listado albaranes por proveedor
            indFrame = 16
            H = 5235
            If OpcionListado = 310 Or OpcionListado = 312 Then
                H = 4325
                Me.cmdAceptarCompras.Top = 3400
                Me.cmdCancel(indFrame).Top = Me.cmdAceptarCompras.Top
                If OpcionListado = 312 Then
                    Me.Label9(21).Caption = "Albaranes por Proveedor"
                Else
                    Me.Label9(21).Caption = "Compras por Proveedor"
                End If
                Me.Label4(87).Caption = "Fecha albaran"
            End If
            W = 7035
            
            PonerFrameVisible Me.FrameCompras, True, H, W
            Me.Frame8.visible = (OpcionListado = 311)
            Me.Frame9.visible = (OpcionListado = 311)
            chkDatosAlbaranes(1).visible = (OpcionListado = 311)
            
    End Select
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
End Sub






Private Sub frmCP_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de cod Postal
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMen_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        If OpcionListado = 305 Or OpcionListado = 306 Then 'Proveedores
            cadFormula = "{sprove.codprove} IN [" & CadenaSeleccion & "]"
            cadSelect = "sprove.codprove IN (" & CadenaSeleccion & ")"
        Else
            If OpcionListado = 190 Or OpcionListado = 191 Then 'Socios
                cadFormula = "{sclien.codclien} IN [" & CadenaSeleccion & "]"
                cadSelect = "sclien.codclien IN (" & CadenaSeleccion & ")"
            Else 'clientes
                cadFormula = "{scliente.codclien} IN [" & CadenaSeleccion & "]"
                cadSelect = "scliente.codclien IN (" & CadenaSeleccion & ")"
            End If
        End If
    Else 'no seleccionamos ningun cliente
        cadFormula = ""
        cadSelect = ""
    End If
End Sub


Private Sub frmMen2_DatoSeleccionado(CadenaSeleccion As String)
'formulario de mensajes: frmMensajes
    If CadenaSeleccion <> "" Then
        cadFormula = cadFormula & " and {sclien.codsitua} IN [" & CadenaSeleccion & "]"
        cadSelect = cadSelect & " and sclien.codsitua IN (" & CadenaSeleccion & ")"
    Else 'no seleccionamos ninguna situacion
        cadFormula = ""
        cadSelect = ""
    End If
End Sub




Private Sub frmMtoActiv_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Actividades
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agentes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArtic_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCartasOfe_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Cartas de Oferta
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFamilia_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Familia de Articulos
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoIncid_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Incidencias
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoProve_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Proveedores
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoSitua_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Situaciones
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSocio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de socio
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtnombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
        Case 0, 1, 39, 40, 45, 60 'Cod. Carta
            Select Case Index
                Case 0: indCodigo = 5
                Case 1: indCodigo = 13
                Case 39: indCodigo = 63
                Case 40: indCodigo = 64
                Case 45: indCodigo = 81
                Case 60: indCodigo = 116
            End Select
            
            Set frmMtoCartasOfe = New frmFacCartasOferta
            frmMtoCartasOfe.DatosADevolverBusqueda = "0|1|"
            frmMtoCartasOfe.Show vbModal
            Set frmMtoCartasOfe = Nothing
            
        Case 2, 3, 4, 5, 9, 10, 24, 46, 47, 56, 57 'Cod. CLIENTE
            Select Case Index
                Case 4, 5: indCodigo = Index - 1
                Case 2, 3: indCodigo = 7 + Index
                Case 9, 10: indCodigo = 18 + Index
                Case 23, 24: indCodigo = Index + 20
                Case 46, 47: indCodigo = Index + 33
                Case 56, 57: indCodigo = Index + 54
            End Select
            Set frmMtoCliente = New frmFacClientes
            frmMtoCliente.DatosADevolverBusqueda = "0|1|"
            frmMtoCliente.Show vbModal
            Set frmMtoCliente = Nothing
        Case 52, 53 'socio
            indCodigo = Index + 44
            Set frmSocio = New frmGesSocios
            frmSocio.DatosADevolverBusqueda = "0|1|"
            frmSocio.Show vbModal
            Set frmSocio = Nothing
        Case 8, 23
            If Index = 8 Then
                indCodigo = 14
            Else
                indCodigo = 15
            End If
            Set frmSocio = New frmGesSocios
            frmSocio.DatosADevolverBusqueda = "0|1|"
            frmSocio.Show vbModal
            Set frmSocio = Nothing
        
        Case 6, 7, 11, 12, 19, 20, 25, 26  'Cod. AGENTE
            Select Case Index
                Case 4, 5: indCodigo = 7 + Index
                Case 5: indCodigo = 12
                Case 6, 7: indCodigo = 12 + Index
                Case 11, 12: indCodigo = 18 + Index
                Case 19, 20, 25, 26: indCodigo = 20 + Index
            End Select
            If OpcionListado <> 92 Then
                Set frmMtoAgente = New frmFacAgentesCom
                frmMtoAgente.DatosADevolverBusqueda = "0|1|"
                frmMtoAgente.Show vbModal
                Set frmMtoAgente = Nothing
            ElseIf Index = 6 Or Index = 7 Then 'Gastos financieros (trabajador)
                Set frmMtoTraba = New frmAdmTrabajadores
                frmMtoTraba.DatosADevolverBusqueda = "0|1|"
                frmMtoTraba.Show vbModal
                Set frmMtoTraba = Nothing
            End If
            
        Case 8, 28, 61, 62 'cod. TRABAJADOR
            indCodigo = 24
            If Index = 28 Then
                indCodigo = 51
            ElseIf Index > 28 Then indCodigo = (117 + 61) - Index
            End If
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 13, 14, 30, 31 'cod. ACTIVIDAD
            indCodigo = 20 + Index
            If Index = 30 Or Index = 31 Then indCodigo = Index + 23
            Set frmMtoActiv = New frmFacActividades
            frmMtoActiv.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtcodigo(indCodigo).Text) Then txtcodigo(indCodigo).Text = ""
            frmMtoActiv.Show vbModal
            Set frmMtoActiv = Nothing
            
           
          
        Case 21, 22, 34 'cod. SITUACION
            indCodigo = 20 + Index
            If Index = 34 Then indCodigo = Index + 23
            Set frmMtoSitua = New frmFacSituaciones
            frmMtoSitua.DatosADevolverBusqueda = "0|1|"
            frmMtoSitua.Show vbModal
            Set frmMtoSitua = Nothing
            
        Case 29 'INCIDENCIAS
            indCodigo = 52
            Set frmMtoIncid = New frmIncidencias
            frmMtoIncid.DatosADevolverBusqueda = "0|1|"
            txtcodigo(indCodigo).Text = ""
            frmMtoIncid.Show vbModal
            Set frmMtoIncid = Nothing
            
        Case 32, 33, 37, 38 'Cod POSTAL
            indCodigo = Index + 23
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0|1|"
            txtcodigo(indCodigo).Text = ""
            frmCP.Show vbModal
            Set frmCP = Nothing
            
        Case 35, 36, 41, 42, 48, 49 'cod. PROVEEDOR
            Select Case Index
                Case 35, 36: indCodigo = Index + 23
                Case 41, 42: indCodigo = Index + 24
                Case 48, 49: indCodigo = Index + 42
            End Select
'            If Index = 35 Or Index = 36 Then indCodigo = Index + 23
'            If Index = 41 Or Index = 42 Then indCodigo = Index + 24
'            If Index = 48 Or Index = 49 Then indCodigo = Index + 42
            Set frmMtoProve = New frmComProveedores
            frmMtoProve.DatosADevolverBusqueda = "0|1|"
            frmMtoProve.Show vbModal
            Set frmMtoProve = Nothing
            
        Case 43, 44, 58, 59 'cod. ARTICULO
            If Index <= 44 Then
                indCodigo = Index + 24
            Else
                indCodigo = Index + 54  'En listado de vetnas x familia articulo
            End If
            Set frmMtoArtic = New frmAlmArticulos
            frmMtoArtic.DatosADevolverBusqueda2 = "@1@" 'Abrimos en Modo Busqueda
            frmMtoArtic.Show vbModal
            Set frmMtoArtic = Nothing
            
        Case 50, 51, 54, 55 'Cod. FAMILIA articulo
            Select Case Index
                Case 50, 51: indCodigo = Index + 44
                Case 54, 55: indCodigo = Index + 46
            End Select
            Set frmMtoFamilia = New frmAlmFamiliaArticulo
            frmMtoFamilia.DatosADevolverBusqueda = "0|1|"
            frmMtoFamilia.Show vbModal
            Set frmMtoFamilia = Nothing
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub imgClearCmbTipomov_Click()
    cboTipomov(2).ListIndex = -1
End Sub

Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 1 'frameOfertas (indFrame=6)
            indCodigo = 3 'Desde
        Case 2 'frameOfertas (indFrame=6)
            indCodigo = 4 'Hasta
        Case 3 'frameRecordatorio Oferta
            indCodigo = 7 '(Desde)
        Case 4 'frameRecordatorio Oferta
            indCodigo = 8 '(Hasta)
        Case 5 'frameEfectuadas
            indCodigo = 16 'Desde
        Case 6 'frameEfectuadas
            indCodigo = 17 'Hasta
        Case 7 'frameTraspasoHco
            indCodigo = 22 'Desde
        Case 8 'frameTraspasoHco
            indCodigo = 23 'hasta
        Case 9, 10 'FrameGenerarPedido
            indCodigo = Index + 16
        Case 11, 12 'Frame Clientes Inactivos
            indCodigo = 20 + Index
        Case 13 'frame pasar pedido a Albaran de compras (a proveedor)
            indCodigo = 49
        Case 14
            indCodigo = 50
        Case 15, 16
            indCodigo = Index + 54
        Case 17 'Frame Factura Rectificariva
            indCodigo = 72
        Case 18, 19 'Ped. Compras
            indCodigo = Index + 56
        Case 20, 21 'Carta Pedidos
            indCodigo = Index + 57
        Case 22: indCodigo = Index + 60
        Case 23, 24 'Reimprimir facturas
            indCodigo = Index + 62
        Case 25, 26 'Cierre caja TPV
            indCodigo = Index + 63
        Case 27, 28 'Listados estadistica compras
            indCodigo = Index + 65
        Case 29, 30 'Estadistica ventas por familia
            indCodigo = Index + 69
   
        Case 31, 32 'Impresion etiq. clientes. Desde / hasta factura
            indCodigo = Index + 73
        Case 33, 34
            indCodigo = Index + 75
   End Select
   
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)
End Sub












Private Sub ListTipoMov_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optEnvioMail_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub optForpago_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 1 Then KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 33 And KeyCode = 38 Then Exit Sub
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Tabla As String
Dim codCampo As String, NomCampo As String
Dim TipCampo As String, Formato As String
Dim Titulo As String
Dim EsNomCod As Boolean


    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    EsNomCod = False
    TipCampo = "N" 'Casi todos son numericos
    
    Select Case Index
        'FECHA Desde Hasta
        Case 1, 2, 8, 16, 17, 22, 23, 25, 26, 31, 32, 49, 50, 69, 70, 72, 74, 75, 77, 78, 82, 85, 86, 88, 89, 92, 93, 98, 99, 104, 105, 108, 109
            If txtcodigo(Index).Text = "" Then Exit Sub
            PonerFormatoFecha txtcodigo(Index)
            
            'Fecha entrega para Pedido. Poner la semana
            If Index = 26 Then
                'Comprobar que fecha entrega es posterior a la del pedido
                If Not EsFechaIgualPosterior(txtcodigo(25).Text, txtcodigo(26).Text, True, "La Fecha de Entrega debe ser posterior a la Fecha del Pedido.") Then
                    txtcodigo(Index).Text = ""
                    PonerFoco txtcodigo(Index)
                Else
                    txtnombre(4).Text = CalculaSemana(CDate(txtcodigo(26).Text))
                End If
            End If
            
        Case 6, 20, 21, 71, 83, 84  'Nº de OFERTA/FACTURA
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
            End If
        
        Case 5, 13, 63, 64, 81, 116 'CARTA de la Oferta
            EsNomCod = True
            Tabla = "scartas"
            codCampo = "codcarta"
            NomCampo = "descarta"
            Formato = "000"
            Titulo = "cod. de Carta"
                    
        Case 9, 10, 43, 44, 79, 80, 96, 97, 110, 111 'Cod. socio
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Socio"


        Case 3, 4, 27, 28 ' codigo cliente
            EsNomCod = True
            Tabla = "scliente"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Cliente"
        
        Case 14, 15 ' codigo socio
            EsNomCod = True
            Tabla = "sclien"
            codCampo = "codclien"
            NomCampo = "nomclien"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Socio"

        Case 11, 12, 18, 19, 29, 30, 39, 40, 45, 46 'Cod. AGENTE
            EsNomCod = True
            Formato = "0000"
            If OpcionListado = 92 Then 'Gastos tecnicos
                If Index = 18 Or Index = 19 Then
                    'cod agente / cod. trabajador
                    Tabla = "straba"
                    codCampo = "codtraba"
                    NomCampo = "nomtraba"
                    Titulo = "Trabajador"
                End If
            Else
                Tabla = "sagent"
                codCampo = "codagent"
                NomCampo = "nomagent"
                Titulo = "Agente"
            End If
        
        Case 24, 47, 51, 117, 118 'Cod. TRABAJADOR
            EsNomCod = True
            Tabla = "straba"
            codCampo = "codtraba"
            NomCampo = "nomtraba"
            Formato = "0000"
            Titulo = "Trabajador"
            
        Case 33, 34, 53, 54 'Cod ACTIVIDAD
            EsNomCod = True
            Tabla = "sactiv"
            codCampo = "codactiv"
            NomCampo = "nomactiv"
            Formato = "000"
            Titulo = "Actividad de Cliente"
            
           
                      
        Case 41, 42, 57 'cod SITUACION
            EsNomCod = True
            Tabla = "ssitua"
            codCampo = "codsitua"
            NomCampo = "nomsitua"
            Formato = "00"
            Titulo = "Situación Especial"
            
        Case 52 'cod. Incidencias
            EsNomCod = True
            Tabla = "sincid"
            codCampo = "codincid"
            NomCampo = "nomincid"
            TipCampo = "T"
            Titulo = "Incidencias"
            
        Case 55, 56, 60, 61 'cod POSTAL
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, "scpostal", "provincia", "cpostal", "CPostal")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = txtcodigo(Index).Text
            
         Case 58, 59, 65, 66, 90, 91 'Cod. PROVEEDOR
            EsNomCod = True
            Tabla = "sprove"
            codCampo = "codprove"
            NomCampo = "nomprove"
            TipCampo = "N"
            Formato = "000000"
            Titulo = "Proveedor"
            
        Case 67, 68, 112, 113 'cod. ARTICULO
            EsNomCod = True
            Tabla = "sartic"
            codCampo = "codartic"
            NomCampo = "nomartic"
            TipCampo = "T"
            Titulo = "Artículo"
            
        Case 73  'Nº de Pedido de Compras
            If txtcodigo(Index).Text = "" Then Exit Sub
            If OpcionListado = 55 Or OpcionListado = 56 Then
                NomCampo = "numpedpr"
                Titulo = "Proveedor"
            Else
                NomCampo = "numpedcl"
                Titulo = "Cliente"
            End If
            NomCampo = DevuelveDesdeBDNew(conAri, NomTabla, NomCampo, NomCampo, txtcodigo(Index).Text, "N")
            If NomCampo = "" Then
                MsgBox "No existe el Nº de Pedido de " & Titulo & ": " & txtcodigo(Index).Text, vbInformation
                txtcodigo(Index).Text = ""
                PonerFoco txtcodigo(Index)
            Else
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
            End If
            
        Case 94, 95, 100, 101 'cod. FAMILIA articulos
            EsNomCod = True
            Tabla = "sfamia"
            codCampo = "codfamia"
            NomCampo = "nomfamia"
            TipCampo = "N"
            Formato = "0000"
            Titulo = "Familia"
    End Select
    
    If EsNomCod Then
        If TipCampo = "N" Then
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, Tabla, NomCampo, codCampo, Titulo, TipCampo)
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, Formato)
            Else
                txtnombre(Index).Text = ""
            End If
        Else
            txtnombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), conAri, Tabla, NomCampo, codCampo, Titulo, TipCampo)
        End If
    End If
End Sub


   
Private Sub PonerFrameClienInacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Clientes Inactivos Visible y Ajustado al Formulario, y visualiza los controles
'necesarios
Dim b As Boolean

    If OpcionListado = 90 Or OpcionListado = 91 Then
        H = 6980
        Me.cmdAceptarClienInac.Top = 5980
        Me.cmdCancel(5).Top = 5980
    Else
        H = 4460
        Me.cmdAceptarClienInac.Top = 3800
        Me.cmdCancel(5).Top = 3800
    End If
    Me.frameCliexFacturas.visible = OpcionListado = 90
    
    If OpcionListado = 90 Or OpcionListado = 91 Then
        W = 11000
    Else
        W = 6800
    End If
    
    PonerFrameVisible Me.FrameClienInactivos, visible, H, W

    If visible = True Then
        b = (OpcionListado = 48)
        'Mostrar D/H Fecha
        Label4(43).visible = b
        Label4(44).visible = b
        Me.imgFecha(12).visible = b
        Me.txtcodigo(32).visible = b
        
        If b Then
            Me.Label4(36).Caption = "Fecha Alta"
            Me.Label8.Caption = "Altas Nuevos Clientes"
        ElseIf OpcionListado = 90 Or OpcionListado = 91 Then
            Me.Frame1.visible = True
            Me.txtcodigo(31).visible = False
            Me.FrameImpClien.visible = True
            Me.OptCliTodos.Value = True
            If OpcionListado = 90 Then
                Me.Label8.Caption = "Etiquetas de Clientes"
                Me.FrameImpClien.Top = 5740
                Me.FrameImpClien.Left = 600
            Else
                Me.Label8.Caption = "Cartas a Clientes"
                Me.FrameImpClien.Left = 6800
                Me.FrameImpClien.Top = 4500
            End If
        End If
        Me.Frame4.visible = (OpcionListado = 91)
    End If
End Sub

Private Sub PonerFrameSociosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Clientes Inactivos Visible y Ajustado al Formulario, y visualiza los controles
'necesarios
Dim b As Boolean

    
    PonerFrameVisible Me.FrameSocios, visible, H, W

End Sub




Private Function AnyadirParametroDH(Cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next
    If txtcodigo(indD).Text <> "" Then
        Cad = Cad & "desde " & txtcodigo(indD).Text
        If txtnombre(indD).Text <> "" Then Cad = Cad & " - " & txtnombre(indD).Text
    End If
    If txtcodigo(indH).Text <> "" Then
        Cad = Cad & "  hasta " & txtcodigo(indH).Text
        If txtnombre(indH).Text <> "" Then Cad = Cad & " - " & txtnombre(indH).Text
    End If
    AnyadirParametroDH = Cad
    If Err.Number <> 0 Then Err.Clear
End Function


Private Function TraspasoOfertaAHco(cadWHERE As String) As Boolean
'Realiza el traspaso de las ofertas seleccionadas por cadWhere
'Inserta en la tabla de Historico de ofertas (schpre, slhpre)
'Borra de las tablas de Ofertas (scapre, slipre)
Dim Sql As String
Dim Donde As String
Dim bol As Boolean

'Aqui empieza transaccion
    conn.BeginTrans
    On Error GoTo ETraspasoHco
    bol = ActualizarElTraspaso(Donde, cadWHERE, "OFE")

ETraspasoHco:
        If Err.Number <> 0 Then
            Sql = "Traspaso Ofertas a Histórico." & vbCrLf & "----------------------------" & vbCrLf
            Sql = Sql & Donde
            MuestraError Err.Number, Sql, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            TraspasoOfertaAHco = True
        Else
            conn.RollbackTrans
            TraspasoOfertaAHco = False
        End If
End Function


Private Function ObtenerTotalOferPeriodo(cadWHERE As String, TotImpA As String, TotImpNA As String, TotOfeA As String, TotOfeNA As String) As Boolean
'para INFORME DE OFERTAS EFECTUADAS
'TotImpA: suma del Importe bruto de todas las Ofertas Aceptadas del periodo seleccionado
'TotImpNA: suma del Importe bruto de todas las Ofertas NO Aceptadas del periodo
'TotOfeA: nº total de ofertas Aceptadas en el periodo
'TotOfeNA: nº total de Ofertas NO Aceptadas en el periodo
Dim Sql As String
Dim RS As ADODB.Recordset
Dim ImpBrutoLin As Currency
Dim ImpBrutoTotA As Currency
Dim ImpBrutoTotNA As Currency
Dim TotalOfeA As Integer
Dim TotalOfeNA As Integer
On Error GoTo ETotalPeriodo

    Sql = "SELECT scapre.numofert, scapre.fecofert,aceptado, dtoppago, dtognral, SUM(importel) as ImpTotal, (sum(importel)*dtoppago)/100 as Impdtopp, (sum(importel)*dtognral)/100 as Impdtogn "
    Sql = Sql & " FROM scapre INNER join slipre ON scapre.numofert=slipre.numofert "
    If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE
    Sql = Sql & " GROUP by scapre.numofert "
    Sql = Sql & " Union "
    Sql = Sql & " SELECT schpre.numofert, schpre.fecofert,aceptado, dtoppago, dtognral, SUM(importel) as ImpTotal,(sum(importel)*dtoppago)/100 as Impdtopp, (sum(importel)*dtognral)/100 as Impdtogn "
    Sql = Sql & " FROM schpre iNNER join slhpre ON schpre.numofert=slhpre.numofert "
    If cadWHERE <> "" Then
'        cadWHERE = SustituirCadenas(cadWHERE, "scapre", "schpre")
        cadWHERE = Replace(cadWHERE, "scapre", "schpre")
        Sql = Sql & " WHERE " & cadWHERE
    End If
    Sql = Sql & " GROUP by schpre.numofert "

    ImpBrutoTotA = 0
    ImpBrutoTotNA = 0
    TotalOfeA = 0
    TotalOfeNA = 0
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        ImpBrutoLin = RS!ImpTotal - RS!impdtopp - RS!impdtogn
        If RS!aceptado = 1 Then 'OFERTA ACEPTADA
            TotalOfeA = TotalOfeA + 1
            ImpBrutoTotA = ImpBrutoTotA + ImpBrutoLin
        Else 'OFERTA NO ACEPTADA
            TotalOfeNA = TotalOfeNA + 1
            ImpBrutoTotNA = ImpBrutoTotNA + ImpBrutoLin
        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    TotImpA = Format(ImpBrutoTotA, "0.00")
    TotImpNA = Format(ImpBrutoTotNA, "0.00")
    TotOfeA = TotalOfeA
    TotOfeNA = TotalOfeNA
    ObtenerTotalOferPeriodo = True
    
ETotalPeriodo:
    If Err.Number <> 0 Then ObtenerTotalOferPeriodo = False
End Function


Private Sub CargarListViewOrden()
'Carga el List View del frame: frameClientes
'para establecer el orden en que se van a mostrar los datos en el Informe
'Orden: Actividad, Zona, Ruta, Agente, Situación
Dim ItmX As ListItem

    'Los encabezados
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Campo", 1500

    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Actividad"
    Set ItmX = ListView1.ListItems.Add
    ItmX.Text = "Agente"
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String
Dim Cad As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    
    'para MySQL
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        'Fecha para la Base de Datos
        Cad = CadenaDesdeHastaBD(txtcodigo(indD).Text, txtcodigo(indH).Text, campo, Tipo)
        If Not AnyadirAFormula(cadSelect, Cad) Then Exit Function
    End If
    
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            cadParam = cadParam & AnyadirParametroDH(param, indD, indH) & """|"
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

'Llevara empipados los siguientes datos, para el envio del mail
'DatosEnvioMail:
'       outTipoDocumento
'       outCodigoCliProv
'       outClaveNombreArchiv
Private Sub LlamarImprimir(PonerNombrePDF As Boolean, EnviaPorEmail As Boolean, Optional DatosEnvioMail As String)
     With frmImprimir
        
        If EnviaPorEmail Then
            If Dir(App.Path & "\docum.pdf") <> "" Then Kill App.Path & "\docum.pdf"
        End If
        
        .outTipoDocumento = 0
        If DatosEnvioMail <> "" Then
            .outTipoDocumento = RecuperaValor(DatosEnvioMail, 1)
            .outCodigoCliProv = RecuperaValor(DatosEnvioMail, 2)
            .outClaveNombreArchiv = RecuperaValor(DatosEnvioMail, 3)
        End If
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = EnviaPorEmail
        .Opcion = OpcionListado
        .Titulo = Titulo
        .NombreRPT = nomRPT
        If PonerNombrePDF Then .NombrePDF = cadPDFrpt
        .ConSubInforme = conSubRPT
        .Show vbModal
    End With
    If DatosEnvioMail <> "" Then DatosEnvioMail = ""
End Sub




Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
'Pone por que campos se van a AGrupar los datos en el Informe de Crystal Report
'El informe tiene definido 4 formulas a las cuales ahora le asignamos un campo
'de la tabla segun el orden seleccionado para el agrupamiento
Dim campo As String
Dim NomCampo As String

    campo = "pGroup" & numGrupo & "="
    NomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0
    
    Select Case cadgrupo
        Case "Actividad"
            cadParam = cadParam & campo & "{scliente.codactiv}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""ACTIVIDAD:  "" & " & " totext({scliente.codactiv},""000"") & " & """  """ & " & {sactiv.nomactiv}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codactiv},""000"") & " & """ """ & " & {sactiv.nomactiv}" & "|"
                cadParam = cadParam & NomCampo & "{sactiv.nomactiv}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Actividad""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
            
            
'            PonerGrupo = numGrupo
        Case "Agente"
            cadParam = cadParam & campo & "{scliente.codagent}" & "|"
            If numGrupo = 1 Then
                cadParam = cadParam & NomCampo & " ""AGENTE:  "" & " & " totext({scliente.codagent},""000000"") & " & """  """ & " & {sagent.nomagent}" & "|"
            Else
'                cadParam = cadParam & nomcampo & " totext({sclien.codagent},""000000"") & " & """ """ & " & {sagent.nomagent}" & "|"
                cadParam = cadParam & NomCampo & "{sagent.nomagent}" & "|"
                cadParam = cadParam & "pTitulo" & numGrupo & "=""Agente""" & "|"
                numParam = numParam + 1
            End If
            numParam = numParam + 1
'        Case "Situacion"
    End Select
End Function


Private Function ListaClientesMante(cadWHERE As String) As String
'devuelve de los clientes filtrados en la cadWhere aquellos que tiene mantenimientos
Dim Sql As String, Cad As String
Dim RS As ADODB.Recordset
On Error GoTo ELista

    Cad = ""
    Sql = "SELECT sclien.codclien "
    Sql = Sql & " FROM sclien INNER JOIN scaman ON sclien.codclien=scaman.codclien "
    If cadWHERE <> "" Then Sql = Sql & " WHERE " & cadWHERE

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & RS.Fields(0).Value & ","
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'quitamos la ultima coma
    Cad = Mid(Cad, 1, Len(Cad) - 1)
    ListaClientesMante = Cad
ELista:
    If Err.Number <> 0 Then MuestraError Err.Number, "Clientes con mantenimientos", Err.Description
End Function




Private Function ListaClientesDesdeHastaFactura2() As String
'devuelve de los clientes filtrados en la cadWhere aquellos que tiene mantenimientos
Dim Sql As String, Cad As String
Dim RS As ADODB.Recordset
On Error GoTo ELista2

    'Monto el cad
    Cad = ""
    If Me.cboTipomov(2).ListIndex >= 0 Then
        'Tipo mov=
        Cad = " AND codtipom = '" & Mid(Me.cboTipomov(2).List(Me.cboTipomov(2).ListIndex), 1, 3) & "'"
    End If
    If txtcodigo(102).Text <> "" Then Cad = Cad & " AND numfactu >= " & txtcodigo(102).Text
    If txtcodigo(103).Text <> "" Then Cad = Cad & " AND numfactu <= " & txtcodigo(103).Text
    If txtcodigo(104).Text <> "" Then Cad = Cad & " AND fecfactu >= '" & Format(txtcodigo(104).Text, FormatoFecha) & "'"
    If txtcodigo(105).Text <> "" Then Cad = Cad & " AND fecfactu <= '" & Format(txtcodigo(105).Text, FormatoFecha) & "'"
    If Len(Cad) > 0 Then Cad = Mid(Cad, 5) 'QUITO EL PRIMER AND
    
    
    
    'Febrero 2010
    'Si no pongo ningun dato para el desde / hasta factura, no me busca en facturados
    If Cad = "" Then
        ListaClientesDesdeHastaFactura2 = ""
        Exit Function
    End If
    
    Sql = "SELECT DISTINCT(scafaccli.codclien) "
    Sql = Sql & " FROM scafaccli "
    If Cad <> "" Then Sql = Sql & " WHERE " & Cad


    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & RS.Fields(0).Value & ","
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'quitamos la ultima coma
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    
    ListaClientesDesdeHastaFactura2 = Cad
ELista2:
    If Err.Number <> 0 Then MuestraError Err.Number, "Procedimiento: ListaClientesDesdeHastaFactura", Err.Description
End Function



Private Sub EnviarEMailMulti(cadWHERE As String, cadTit As String, cadRpt As String, cadTabla As String)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim cad1 As String, cad2 As String, Lista As String
Dim cont As Integer

On Error GoTo EEnviar

    Screen.MousePointer = vbHourglass
    
    If cadTabla = "sprove" Then
        'seleccionamos todos los proveedores a los que queremos enviar e-mail
        Sql = "SELECT codprove,nomprove,maiprov1,maiprov2 "
    ElseIf cadTabla = "sclien" Then
        'seleccionamos todos los clientes a los que queremos enviar e-mail
        Sql = "SELECT codclien,nomclien,maiclie1,maiclie2 "
    End If
    Sql = Sql & "FROM " & cadTabla
    Sql = Sql & " WHERE " & cadWHERE
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'creamos una temporal donde guardamos para cada proveedor que SI tiene
    'e-mail, el mail1 o el mail2 al que vamos a enviar
    Sql = "CREATE TEMPORARY TABLE tmpMail ( "
    Sql = Sql & "codusu SMALLINT(3) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "codprove INT(6) UNSIGNED  DEFAULT '0' NOT NULL, "
    Sql = Sql & "nomprove varchar(40)  DEFAULT '' NOT NULL, "
    Sql = Sql & "email varchar(40)  DEFAULT '' NOT NULL) "
    conn.Execute Sql
    
    cont = 0
    Lista = ""
    
    While Not RS.EOF
    'para cada cliente/proveedor enviamos un e-mail
        cad1 = DBLet(RS.Fields(2), "T") 'e-mail administracion
        cad2 = DBLet(RS.Fields(3), "T") 'e-mail compras
        
        If cad1 = "" And cad2 = "" Then 'no tiene e-mail
'              MsgBox "Sin mail para el proveedor: " & Format(RS!codProve, "000000") & " - " & RS!nomprove, vbExclamation
              Lista = Lista & Format(RS.Fields(0), "000000") & " - " & RS.Fields(1) & vbCrLf
        ElseIf cad1 <> "" And cad2 <> "" Then 'tiene 2 e-mail
            'ver a q e-mail se va a enviar (administracion, compras)
            If cadTabla = "sprove" Then
                If Me.OptMailCom(0).Value = True Then cad1 = cad2
            Else
                If Me.OptMailCom(1).Value = True Then cad1 = cad2
            End If
        Else 'alguno de los 2 tiene valor
            If cad2 <> "" Then cad1 = cad2  'e-mail para compras
        End If
        
        If cad1 <> "" Then 'HAY email --> ENVIAMOS e-mail
            With frmImprimir
                .OtrosParametros = cadParam
                .NumeroParametros = numParam
                If cadTabla = "sprove" Then
                    Sql = "{sprove.codprove}=" & RS.Fields(0)
                    .Opcion = 306
                Else
                    Sql = "{sclien.codclien}=" & RS.Fields(0)
                    .Opcion = 91
                End If
                .FormulaSeleccion = Sql
                .EnvioEMail = True
                CadenaDesdeOtroForm = "GENERANDO"
                .Titulo = cadTit
                .NombreRPT = cadRpt
                .ConSubInforme = True
                .Show vbModal

                If CadenaDesdeOtroForm = "" Then
                'si se ha generado el .pdf para enviar
                    Sql = "INSERT INTO tmpMail (codusu,codprove,nomprove,email)"
                    Sql = Sql & " VALUES (" & vUsu.Codigo & "," & DBSet(RS.Fields(0), "N") & "," & DBSet(RS.Fields(1), "T") & "," & DBSet(cad1, "T") & ")"
                    conn.Execute Sql
            
                    Me.Refresh
                    Espera 0.4
                    cont = cont + 1
                    'Se ha generado bien el documento
                    'Lo copiamos sobre app.path & \temp
                    Sql = RS.Fields(0) & ".pdf"
                    FileCopy App.Path & "\docum.pdf", App.Path & "\temp\" & Sql
                End If
            End With
        End If
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
      
    If cont > 0 Then
        Espera 0.4
        If cadTabla = "sprove" Then
            Sql = "Carta: " & txtnombre(63).Text & "|"
             Sql = Sql & "Att : " & txtcodigo(62).Text & "|"
        Else
            Sql = "Carta: " & txtnombre(64).Text & "|"
            Sql = Sql & "Att : " & txtcodigo(0).Text & "|"
        End If
       
        frmEMail.Opcion = 2
        frmEMail.DatosEnvio = Sql
        frmEMail.Show vbModal

        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute Sql
        
        'Borrar la carpeta con temporales
        Kill App.Path & "\temp\*.pdf"
    End If
    
    Screen.MousePointer = vbDefault
   
    'Mostra mensaje con aquellos proveedores que no tienen e-mail
    If Lista <> "" Then
        If cadTabla = "sprove" Then
            Lista = "Proveedores sin e-mail:" & vbCrLf & vbCrLf & Lista
        Else
            Lista = "Clientes sin e-mail:" & vbCrLf & vbCrLf & Lista
        End If
        MsgBox Lista, vbInformation
    End If
    
EEnviar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Enviando Informe por e-mail", Err.Description
        'Borrar la tabla temporal
        Sql = " DROP TABLE IF EXISTS tmpMail;"
        conn.Execute Sql
    End If
End Sub




Private Sub CargarComboTipoMov(indice As Integer)
'### Combo Tipo Movimiento
'Cargaremos el combo, o bien desde una tabla o con valores fijos
' El estilo del combo debe de ser 2 - Dropdown List
' Si queremos que este ordenado, o lo ordenamos por la sentencia sql
' o marcamos la opcion sorted del combo

'Lo cargamos con los valores de la tabla stipom que tengan tipo de documento=Albaranes (tipodocu=1)
Dim Sql As String
Dim RS As ADODB.Recordset
Dim i As Byte

    On Error GoTo ECargaCombo

'    SQL = "select codtipom, nomtipom from stipom where tipodocu=2 " 'Documentos de Facturas
    '3 abril 2007.
    'Mostraba todas las facturas (movimientos que empizan por F, excepto las rectificativas
    'AHora tiene que mostrarlas todas
    'SQL = "select codtipom, nomtipom from stipom where (codtipom like 'F__') and (codtipom<>'FRT')"
    Sql = "select codtipom, nomtipom from stipom where (codtipom like 'F__')"  ' and (codtipom<>'FRT')"
    
    If NumCod <> "" Then Sql = Sql & " and " & NumCod
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    
    If indice < 1000 Then
            'Son combos normales
         cboTipomov(indice).Clear
        
         While Not RS.EOF
             cboTipomov(indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
             cboTipomov(indice).ItemData(cboTipomov(indice).NewIndex) = i
             i = i + 1
             RS.MoveNext
         Wend
        
    
    Else
        
'        ListTipoMov(Indice).Clear
        
        'LOS TIKCETS NO LOS ENVIO POR MAIL
        While Not RS.EOF
            If RS!codtipom <> "FTI" Then
            
'                ListTipoMov(Indice).AddItem RS.Fields(0).Value & "-" & RS.Fields(1).Value
'                'ListTipoMov(indice).List (ListTipoMov(indice).NewIndex)
'                ListTipoMov(Indice).Selected((ListTipoMov(Indice).NewIndex)) = True
            End If
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing
    
    'Pongo el dos para todos menos para la de etiquetas cliente
    If indice < 1000 Then
        If indice <> 2 Then Me.cboTipomov(indice).ListIndex = 2
    End If
ECargaCombo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub PonerFramePedVisible(H As Integer, W As Integer)
'Frame de Pedidos de Venta y Compra
    W = 6075
    H = 4455
    PonerFrameVisible Me.FramePedidos, True, H, W
    Select Case OpcionListado
        Case 38 'PEdidos venta
            Me.Label12(0).Caption = "Informe Pedidos ventas"
            NomTabla = "scaped"
            NomTablaLin = "sliped"
            Me.Label12(3).Caption = "Imprimir otros Pedidos del Cliente:"
        Case 239 'Historico de Pedidos Venta
            Me.Label12(0).Caption = "Informe Hist. Pedidos ventas"
            NomTabla = "schped" 'Cabecera  Hco de Pedidos de clientes
            NomTablaLin = "slhped"
            If FecEntre <> "" Then txtcodigo(76).Text = FecEntre
        Case 55 'Cabecera de Pedidos de Compras (a proveedores)
            Me.Label12(0).Caption = "Informe Pedidos compras"
            NomTabla = "scappr"
            NomTablaLin = "slippr"
        Case 56 'Historico de Pedidos Compras
            Me.Label12(0).Caption = "Informe Hist. Pedidos compras"
            NomTabla = "schppr" 'Cabecera  Hco de Pedidos de Compras (a proveedores)
            NomTablaLin = "slhppr"
            If FecEntre <> "" Then txtcodigo(76).Text = FecEntre
    End Select
    
    
    'Ver Fecha Pedido (En Hist.)
    Label12(2).visible = (OpcionListado = 239) Or OpcionListado = 56
    txtcodigo(76).visible = (OpcionListado = 239) Or OpcionListado = 56
End Sub

 
Private Sub PonerTamnyosMail(peque As Boolean)
    If peque Then
        Me.Height = Me.FrameEnvioMail.Height + 60
        Me.Width = Me.FrameEnvioMail.Width
    Else
'        Me.Height = Me.FrameEnvioFacMail.Height
'        Me.Width = Me.FrameEnvioFacMail.Width
    End If
    Me.Height = Me.Height + 420
    Me.Width = Me.Width + 120
    Me.FrameEnvioMail.visible = peque
    DoEvents
    Me.Refresh
End Sub


Private Sub CargarCombo()
Dim Sql As String
Dim i As Integer
Dim Cad As String
Dim RS As ADODB.Recordset



    Sql = "select codtipom, nomtipom from stipom where codtipom in ('FAV','FTI','FCE','FCN','FAR','FRT','FRC','FTI') "

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Combo1.Clear
    
    i = 0
    Cad = "Todos"
    Combo1.AddItem Cad
    Combo1.ItemData(Combo1.NewIndex) = i
    
    While Not RS.EOF
        i = i + 1
        Cad = RS.Fields(0).Value & " - " & RS.Fields(1).Value
        Combo1.AddItem Cad
        Combo1.ItemData(Combo1.NewIndex) = i
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing

End Sub

Private Sub CargarCombo2()
Dim Sql As String
Dim i As Integer
Dim Cad As String
Dim RS As ADODB.Recordset



    Sql = "select codtipom, nomtipom from stipom where codtipom in ('FAC','FRT','FPC') "

    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Combo2.Clear
    
    i = 0
    Cad = "Todos"
    Combo2.AddItem Cad
    Combo2.ItemData(Combo2.NewIndex) = i
    
    While Not RS.EOF
        i = i + 1
        Cad = RS.Fields(0).Value & " - " & RS.Fields(1).Value
        Combo2.AddItem Cad
        Combo2.ItemData(Combo2.NewIndex) = i
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing

End Sub


