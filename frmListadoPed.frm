VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoPed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   11055
   Icon            =   "frmListadoPed.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FramePreFacturar 
      Height          =   5775
      Left            =   240
      TabIndex        =   40
      Top             =   720
      Width           =   7035
      Begin VB.ComboBox cmbTipAlbaran 
         BeginProperty Font 
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
         ItemData        =   "frmListadoPed.frx":000C
         Left            =   1920
         List            =   "frmListadoPed.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkResumenForpa 
         Caption         =   "Resumen forma de pago"
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
         Left            =   4110
         TabIndex        =   55
         Top             =   4485
         Width           =   2775
      End
      Begin VB.CheckBox chkSoloFacturar 
         Caption         =   "Solo Albaranes para facturar"
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
         Left            =   870
         TabIndex        =   54
         Top             =   4485
         Value           =   1  'Checked
         Width           =   3195
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tipo Informe"
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
         Height          =   735
         Left            =   480
         TabIndex        =   124
         Top             =   3720
         Width           =   5655
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Facturacion"
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
            Left            =   3720
            TabIndex        =   53
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Resumen"
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
            Left            =   2160
            TabIndex        =   52
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle"
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
            Left            =   360
            TabIndex        =   51
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   480
         TabIndex        =   118
         Top             =   2670
         Width           =   6135
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
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   50
            Top             =   600
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
            Index           =   33
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   120
            Text            =   "Text5"
            Top             =   600
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
            Index           =   32
            Left            =   1260
            MaxLength       =   6
            TabIndex        =   49
            Top             =   240
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
            Index           =   32
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   119
            Text            =   "Text5"
            Top             =   240
            Width           =   3615
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
            Index           =   17
            Left            =   330
            TabIndex        =   123
            Top             =   600
            Width           =   570
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
            Index           =   0
            Left            =   330
            TabIndex        =   122
            Top             =   240
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   19
            Left            =   960
            Top             =   615
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
            Index           =   38
            Left            =   0
            TabIndex        =   121
            Top             =   0
            Width           =   765
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   18
            Left            =   960
            Top             =   255
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
         Index           =   26
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarPreFac 
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
         Left            =   4080
         TabIndex        =   57
         Top             =   5040
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
         Index           =   5
         Left            =   5160
         TabIndex        =   58
         Top             =   5040
         Width           =   975
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
         Index           =   27
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   44
         Top             =   1200
         Width           =   1215
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text5"
         Top             =   2880
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
         Index           =   30
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   47
         Top             =   2880
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
         Index           =   31
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text5"
         Top             =   3240
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
         Index           =   31
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   48
         Top             =   3240
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
         Index           =   29
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   46
         Top             =   2280
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
         Index           =   29
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text5"
         Top             =   2280
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
         TabIndex        =   45
         Top             =   1920
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
         Index           =   28
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   1920
         Width           =   3615
      End
      Begin VB.Label lblTipAlbaran 
         Caption         =   "Tipo de albaranes:"
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
         Left            =   480
         TabIndex        =   168
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   3030
         TabIndex        =   70
         Top             =   1200
         Width           =   570
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1440
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Prefacturación de Albaranes"
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
         Left            =   360
         TabIndex        =   69
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Index           =   43
         Left            =   480
         TabIndex        =   68
         Top             =   960
         Width           =   1515
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
         Index           =   42
         Left            =   735
         TabIndex        =   67
         Top             =   1200
         Width           =   600
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3600
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   16
         Left            =   1440
         Top             =   2895
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Formas de Pago"
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
         TabIndex        =   66
         Top             =   2640
         Width           =   1740
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
         Index           =   40
         Left            =   765
         TabIndex        =   65
         Top             =   2880
         Width           =   600
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   17
         Left            =   1440
         Top             =   3255
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
         Index           =   39
         Left            =   765
         TabIndex        =   64
         Top             =   3240
         Width           =   570
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
         Index           =   35
         Left            =   765
         TabIndex        =   63
         Top             =   2250
         Width           =   570
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1440
         Top             =   2280
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
         Index           =   34
         Left            =   765
         TabIndex        =   62
         Top             =   1920
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
         Index           =   33
         Left            =   480
         TabIndex        =   61
         Top             =   1680
         Width           =   555
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1440
         Top             =   1920
         Width           =   240
      End
   End
   Begin VB.Frame FrameEstVentas 
      Height          =   3975
      Left            =   480
      TabIndex        =   157
      Top             =   720
      Width           =   7035
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
         Left            =   1500
         MaxLength       =   4
         TabIndex        =   161
         Top             =   1440
         Width           =   840
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
         Left            =   4200
         TabIndex        =   163
         Top             =   3120
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
         Index           =   8
         Left            =   5280
         TabIndex        =   164
         Top             =   3120
         Width           =   975
      End
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         TabIndex        =   158
         Top             =   1830
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
            Index           =   8
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   159
            Text            =   "Text5"
            Top             =   120
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
            Index           =   8
            Left            =   1020
            MaxLength       =   6
            TabIndex        =   162
            Top             =   120
            Width           =   840
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   32
            Left            =   705
            Top             =   120
            Width           =   240
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
            Index           =   30
            Left            =   0
            TabIndex        =   160
            Top             =   120
            Width           =   555
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Ventas por meses"
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
         Left            =   480
         TabIndex        =   166
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Año"
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
         Left            =   480
         TabIndex        =   165
         Top             =   1440
         Width           =   405
      End
   End
   Begin VB.Frame FrameGenAlbaran 
      Height          =   5895
      Left            =   720
      TabIndex        =   30
      Top             =   480
      Width           =   6675
      Begin VB.CheckBox chkImpHojaExped 
         Caption         =   "Imprimir Hoja Expedición"
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
         Left            =   3810
         TabIndex        =   181
         Top             =   4400
         Width           =   2775
      End
      Begin VB.CheckBox chkImpEtiq 
         Caption         =   "Imprimir Etiquetas"
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
         Left            =   3810
         TabIndex        =   180
         Top             =   4120
         Width           =   2145
      End
      Begin VB.Frame FramepedidoFactura 
         BorderStyle     =   0  'None
         Caption         =   "Frame15"
         Height          =   615
         Left            =   420
         TabIndex        =   171
         Top             =   4560
         Width           =   5985
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
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   172
            Text            =   "Text5"
            Top             =   240
            Width           =   4665
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
            Left            =   375
            MaxLength       =   6
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   0
            Left            =   60
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cta prevista cobro"
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
            Left            =   60
            TabIndex        =   173
            Top             =   -30
            Width           =   1845
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
         Index           =   25
         Left            =   2340
         MaxLength       =   10
         TabIndex        =   20
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CheckBox chkImpAlbaran 
         Caption         =   "Imprimir Albaran"
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
         Left            =   3810
         TabIndex        =   21
         Top             =   3840
         Width           =   2145
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
         Index           =   19
         Left            =   840
         MaxLength       =   4
         TabIndex        =   19
         Top             =   3360
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
         Index           =   19
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text5"
         Top             =   3360
         Width           =   4845
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
         Index           =   18
         Left            =   840
         MaxLength       =   4
         TabIndex        =   18
         Top             =   2640
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
         Index           =   18
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   2640
         Width           =   4845
      End
      Begin VB.CommandButton cmdAceptarGenAlb 
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
         Left            =   4110
         TabIndex        =   23
         Top             =   5280
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
         Index           =   3
         Left            =   5310
         TabIndex        =   24
         Top             =   5280
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
         Index           =   17
         Left            =   840
         MaxLength       =   4
         TabIndex        =   17
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
         Index           =   17
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "Text5"
         Top             =   1920
         Width           =   4845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
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
         Index           =   32
         Left            =   540
         TabIndex        =   39
         Top             =   3840
         Width           =   1410
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   2040
         Top             =   3840
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Envío"
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
         Left            =   540
         TabIndex        =   38
         Top             =   3090
         Width           =   1515
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   10
         Left            =   540
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Material Preparado por"
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
         Left            =   540
         TabIndex        =   36
         Top             =   2370
         Width           =   2235
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   9
         Left            =   540
         Top             =   2640
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Pasar Pedido a "
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
         TabIndex        =   34
         Top             =   480
         Width           =   5685
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Introduzca los siguiente campos: "
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
         Index           =   14
         Left            =   540
         TabIndex        =   33
         Top             =   1200
         Width           =   3675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador de Albaran"
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
         Left            =   540
         TabIndex        =   32
         Top             =   1650
         Width           =   2190
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   8
         Left            =   540
         Top             =   1920
         Width           =   240
      End
   End
   Begin VB.Frame FrameFacturar 
      Height          =   7575
      Left            =   120
      TabIndex        =   71
      Top             =   0
      Width           =   7395
      Begin VB.CommandButton cmdAceptarFacCli 
         Caption         =   "&Aceptar"
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
         Height          =   375
         Left            =   5160
         TabIndex        =   204
         Top             =   6840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame15 
         Height          =   1335
         Left            =   180
         TabIndex        =   175
         Top             =   5160
         Width           =   7005
         Begin VB.TextBox txtCSB 
            BeginProperty Font 
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
            Left            =   2280
            TabIndex        =   87
            Text            =   "Text1"
            Top             =   900
            Width           =   4455
         End
         Begin VB.TextBox txtCSB 
            BeginProperty Font 
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
            Left            =   2280
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   540
            Width           =   4455
         End
         Begin VB.TextBox txtCSB 
            BeginProperty Font 
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
            Left            =   2280
            TabIndex        =   85
            Text            =   "Text1"
            Top             =   180
            Width           =   4455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Texto csb4"
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
            Left            =   1170
            TabIndex        =   179
            Top             =   900
            Width           =   1110
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Texto csb3"
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
            Index           =   25
            Left            =   1170
            TabIndex        =   178
            Top             =   540
            Width           =   1110
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Texto csb2"
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
            Left            =   1170
            TabIndex        =   177
            Top             =   180
            Width           =   1110
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Tesoreria"
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
            Left            =   120
            TabIndex        =   176
            Top             =   180
            Width           =   1020
         End
      End
      Begin VB.Frame FrameProgress 
         Height          =   1050
         Left            =   180
         TabIndex        =   153
         Top             =   6450
         Visible         =   0   'False
         Width           =   4695
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   345
            Left            =   120
            TabIndex        =   154
            Top             =   600
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   156
            Top             =   350
            Width           =   4335
         End
         Begin VB.Label lblProgess 
            Caption         =   "Facturando:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   155
            Top             =   135
            Width           =   4215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3375
         Left            =   180
         TabIndex        =   94
         Top             =   1800
         Width           =   7005
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
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   108
            Text            =   "Text5"
            Top             =   2580
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
            Index           =   42
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   83
            Top             =   2580
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
            Index           =   43
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   107
            Text            =   "Text5"
            Top             =   2970
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
            Index           =   43
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   84
            Top             =   2970
            Width           =   735
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
            Index           =   41
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   82
            Top             =   2040
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
            Index           =   41
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   103
            Text            =   "Text5"
            Top             =   2040
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
            Index           =   40
            Left            =   2220
            MaxLength       =   6
            TabIndex        =   81
            Top             =   1650
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
            Index           =   40
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   102
            Text            =   "Text5"
            Top             =   1650
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
            Index           =   38
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   79
            Top             =   1200
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
            Index           =   39
            Left            =   5400
            MaxLength       =   10
            TabIndex        =   80
            Top             =   1200
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
            Index           =   36
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   77
            Top             =   720
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
            Index           =   37
            Left            =   5400
            MaxLength       =   10
            TabIndex        =   78
            Top             =   720
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
            Index           =   35
            Left            =   3540
            MaxLength       =   10
            TabIndex        =   76
            Top             =   240
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   1920
            Top             =   2580
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
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
            Left            =   240
            TabIndex        =   111
            Top             =   2280
            Width           =   1290
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
            Index           =   48
            Left            =   1215
            TabIndex        =   110
            Top             =   2580
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   23
            Left            =   1920
            Top             =   2970
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
            Index           =   49
            Left            =   1215
            TabIndex        =   109
            Top             =   2970
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   1920
            Top             =   1650
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   1920
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
            Index           =   50
            Left            =   1215
            TabIndex        =   106
            Top             =   2040
            Width           =   570
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
            Index           =   51
            Left            =   1215
            TabIndex        =   105
            Top             =   1650
            Width           =   600
         End
         Begin VB.Label Label10 
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
            Index           =   2
            Left            =   240
            TabIndex        =   104
            Top             =   1530
            Width           =   765
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
            Left            =   4470
            TabIndex        =   101
            Top             =   1200
            Width           =   570
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   2700
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Albaran"
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
            Left            =   240
            TabIndex        =   100
            Top             =   1200
            Width           =   1515
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
            Index           =   46
            Left            =   2025
            TabIndex        =   99
            Top             =   1200
            Width           =   600
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   5100
            Top             =   1215
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
            Index           =   36
            Left            =   4470
            TabIndex        =   98
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Albaran"
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
            Left            =   240
            TabIndex        =   97
            Top             =   720
            Width           =   1155
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
            Index           =   45
            Left            =   2025
            TabIndex        =   96
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Periodicidad de la Facturación"
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
            Left            =   240
            TabIndex        =   95
            Top             =   240
            Width           =   3210
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   180
         TabIndex        =   90
         Top             =   720
         Width           =   7005
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
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   74
            Top             =   210
            Width           =   1335
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
            Left            =   2610
            MaxLength       =   10
            TabIndex        =   73
            Top             =   210
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
            Index           =   0
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   91
            Text            =   "Text5"
            Top             =   600
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
            Index           =   0
            Left            =   2340
            MaxLength       =   6
            TabIndex        =   75
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Factura"
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
            Index           =   9
            Left            =   4200
            TabIndex        =   167
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Facturación"
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
            Left            =   240
            TabIndex        =   93
            Top             =   240
            Width           =   1935
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   2340
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Cta.Prevista Cobro"
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
            Left            =   240
            TabIndex        =   92
            Top             =   600
            Width           =   2055
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   24
            Left            =   1920
            Top             =   600
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdAceptarFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5160
         TabIndex        =   88
         Top             =   6840
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
         Index           =   6
         Left            =   6240
         TabIndex        =   89
         Top             =   6840
         Width           =   975
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   360
         TabIndex        =   174
         Top             =   3360
         Width           =   6615
      End
      Begin VB.Label Label10 
         Caption         =   "Facturación de Albaranes"
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
         Left            =   360
         TabIndex        =   72
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame FramePedxArtic 
      Height          =   7575
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   10635
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   300
         TabIndex        =   137
         Top             =   2580
         Width           =   6015
         Begin VB.Frame Frame11 
            Caption         =   " Ordenar por "
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
            Height          =   1215
            Left            =   60
            TabIndex        =   140
            Top             =   600
            Width           =   2415
            Begin VB.OptionButton OptOrdenVentas 
               Caption         =   "Volumen ventas"
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
               TabIndex        =   143
               Top             =   840
               Value           =   -1  'True
               Width           =   2235
            End
            Begin VB.OptionButton OptOrdenNomclien 
               Caption         =   "Nombre socio"
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
               Left            =   120
               TabIndex        =   142
               Top             =   480
               Width           =   2235
            End
            Begin VB.OptionButton OptOrdenCodclien 
               Caption         =   "Cod. socio"
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
               TabIndex        =   141
               Top             =   240
               Width           =   2115
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
            Index           =   1
            Left            =   4350
            MaxLength       =   15
            TabIndex        =   12
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "€"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   5640
            TabIndex        =   139
            Top             =   260
            Width           =   105
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Mostrar Clientes con ventas superior a"
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
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   4185
         End
      End
      Begin VB.Frame FramepedxClien 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   6570
         TabIndex        =   185
         Top             =   4860
         Visible         =   0   'False
         Width           =   6375
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
            Index           =   10
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   11
            Top             =   1680
            Visible         =   0   'False
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
            Index           =   10
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   193
            Text            =   "Text5"
            Top             =   1680
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
            Index           =   9
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   10
            Top             =   1320
            Visible         =   0   'False
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
            Index           =   9
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   191
            Text            =   "Text5"
            Top             =   1320
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
            Index           =   7
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   9
            Top             =   600
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
            Index           =   7
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   189
            Text            =   "Text5"
            Top             =   600
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
            Index           =   6
            Left            =   1350
            MaxLength       =   6
            TabIndex        =   8
            Top             =   240
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
            Index           =   6
            Left            =   2130
            Locked          =   -1  'True
            TabIndex        =   186
            Text            =   "Text5"
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label Label4 
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
            Height          =   195
            Index           =   52
            Left            =   0
            TabIndex        =   195
            Top             =   1080
            Visible         =   0   'False
            Width           =   615
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
            Index           =   47
            Left            =   300
            TabIndex        =   194
            Top             =   1680
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   34
            Left            =   1050
            Top             =   1680
            Visible         =   0   'False
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
            Index           =   31
            Left            =   300
            TabIndex        =   192
            Top             =   1320
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   33
            Left            =   1050
            Top             =   1320
            Visible         =   0   'False
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
            Index           =   29
            Left            =   300
            TabIndex        =   190
            Top             =   600
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   13
            Left            =   1050
            Top             =   600
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
            Index           =   28
            Left            =   0
            TabIndex        =   188
            Top             =   0
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
            Index           =   27
            Left            =   300
            TabIndex        =   187
            Top             =   240
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   3
            Left            =   1050
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.ComboBox cmbTipAlbaran 
         BeginProperty Font 
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
         ItemData        =   "frmListadoPed.frx":004C
         Left            =   2640
         List            =   "frmListadoPed.frx":0059
         Style           =   2  'Dropdown List
         TabIndex        =   169
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   7080
         TabIndex        =   144
         Top             =   1440
         Width           =   6495
         Begin VB.Frame Frame13 
            Height          =   615
            Left            =   60
            TabIndex        =   150
            Top             =   1320
            Width           =   3045
            Begin VB.OptionButton OptResumen 
               Caption         =   "Resumen"
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
               Left            =   1320
               TabIndex        =   152
               Top             =   240
               Width           =   1365
            End
            Begin VB.OptionButton OptDetalle 
               Caption         =   "Detalle"
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
               TabIndex        =   151
               Top             =   240
               Value           =   -1  'True
               Width           =   1245
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
            Height          =   285
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   146
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
            Height          =   285
            Index           =   2
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   13
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
            Height          =   285
            Index           =   3
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   145
            Text            =   "Text5"
            Top             =   720
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
            Height          =   285
            Index           =   3
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   1
            Left            =   1080
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Trabajador"
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
            Left            =   120
            TabIndex        =   149
            Top             =   120
            Width           =   1185
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
            Index           =   21
            Left            =   420
            TabIndex        =   148
            Top             =   360
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   2
            Left            =   1080
            Top             =   720
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
            Index           =   20
            Left            =   420
            TabIndex        =   147
            Top             =   720
            Width           =   570
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
         Index           =   12
         Left            =   4140
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1470
         Width           =   1215
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
         Index           =   2
         Left            =   5760
         TabIndex        =   16
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptarPedxArtic 
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
         Left            =   4680
         TabIndex        =   15
         Top             =   7080
         Width           =   975
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
         Index           =   11
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   0
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Frame FrameOrden2 
         Height          =   615
         Left            =   240
         TabIndex        =   182
         Top             =   6840
         Width           =   2505
         Begin VB.OptionButton optOrdePed 
            Caption         =   "Pedido"
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
            Left            =   1380
            TabIndex        =   184
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optOrdePed 
            Caption         =   "Articulo"
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
            TabIndex        =   183
            Top             =   240
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   330
         TabIndex        =   125
         Top             =   2580
         Width           =   6975
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "Text5"
            Top             =   480
            Width           =   4215
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
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   4
            Top             =   480
            Width           =   1215
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
            Index           =   16
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "Text5"
            Top             =   840
            Width           =   4215
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
            Index           =   16
            Left            =   1260
            MaxLength       =   16
            TabIndex        =   5
            Top             =   840
            Width           =   1215
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   6
            Left            =   960
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label4 
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
            Index           =   1
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   810
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
            Index           =   3
            Left            =   300
            TabIndex        =   129
            Top             =   480
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   7
            Left            =   960
            Top             =   840
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
            Index           =   9
            Left            =   300
            TabIndex        =   128
            Top             =   840
            Width           =   570
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   112
         Top             =   2550
         Width           =   6495
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
            Index           =   21
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   7
            Top             =   720
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
            Index           =   21
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   114
            Text            =   "Text5"
            Top             =   720
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
            Index           =   20
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   6
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
            Index           =   20
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   113
            Text            =   "Text5"
            Top             =   360
            Width           =   3855
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
            Index           =   12
            Left            =   300
            TabIndex        =   117
            Top             =   720
            Width           =   570
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   12
            Left            =   1080
            Top             =   720
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
            Index           =   13
            Left            =   300
            TabIndex        =   116
            Top             =   360
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
            Left            =   120
            TabIndex        =   115
            Top             =   120
            Width           =   555
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   11
            Left            =   1080
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   196
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
            Index           =   22
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   202
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
            Index           =   22
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   197
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
            Index           =   23
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   203
            Text            =   "Text5"
            Top             =   720
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
            Index           =   23
            Left            =   1380
            MaxLength       =   6
            TabIndex        =   198
            Top             =   720
            Width           =   735
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   26
            Left            =   1080
            Top             =   750
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
            Index           =   55
            Left            =   120
            TabIndex        =   201
            Top             =   120
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
            Index           =   54
            Left            =   300
            TabIndex        =   200
            Top             =   360
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   25
            Left            =   1080
            Top             =   390
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
            Index           =   53
            Left            =   300
            TabIndex        =   199
            Top             =   720
            Width           =   570
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   360
         TabIndex        =   131
         Top             =   1770
         Width           =   6375
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
            Index           =   13
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   133
            Text            =   "Text5"
            Top             =   360
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
            Index           =   13
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   2
            Top             =   360
            Width           =   495
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   132
            Text            =   "Text5"
            Top             =   720
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
            Index           =   14
            Left            =   1260
            MaxLength       =   3
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   270
            Index           =   4
            Left            =   960
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Almacen"
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
            Left            =   120
            TabIndex        =   136
            Top             =   120
            Width           =   915
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
            Height          =   270
            Index           =   7
            Left            =   300
            TabIndex        =   135
            Top             =   360
            Width           =   600
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   5
            Left            =   960
            Top             =   720
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
            Index           =   8
            Left            =   300
            TabIndex        =   134
            Top             =   720
            Width           =   570
         End
      End
      Begin VB.Label lblTipAlbaran 
         Caption         =   "Tipo de albaranes:"
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
         Left            =   600
         TabIndex        =   170
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   7
         Left            =   3840
         Top             =   1470
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
         Index           =   5
         Left            =   660
         TabIndex        =   29
         Top             =   1470
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pedido"
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
         TabIndex        =   28
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label1 
         Caption         =   "Pedidos por Artículo"
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
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   4815
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1440
         Top             =   1470
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
         Index           =   2
         Left            =   3180
         TabIndex        =   26
         Top             =   1470
         Width           =   570
      End
   End
End
Attribute VB_Name = "frmListadoPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)
      
      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir


Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor

'#Laura 14/11/2006 Recuperar facturas Alzira
Public EstaRecupFact As Boolean ' si esta recuperando facturas (para albaranes de mostrador)

Public Escliente As Boolean

'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmMtoCliente As frmFacClientes
Attribute frmMtoCliente.VB_VarHelpID = -1
Private WithEvents frmMtoAlmacen As frmAlmAlPropios
Attribute frmMtoAlmacen.VB_VarHelpID = -1
Private WithEvents frmMtoArticulo As frmAlmArticulos
Attribute frmMtoArticulo.VB_VarHelpID = -1
Private WithEvents frmMtoTraba As frmAdmTrabajadores
Attribute frmMtoTraba.VB_VarHelpID = -1
Private WithEvents frmMtoFEnvio As frmFacFormasEnvio
Attribute frmMtoFEnvio.VB_VarHelpID = -1
Private WithEvents frmMtoFPago As frmFacFormasPago
Attribute frmMtoFPago.VB_VarHelpID = -1
Private WithEvents frmMtoAgente As frmFacAgentesCom
Attribute frmMtoAgente.VB_VarHelpID = -1
Private WithEvents frmMtoBancosPro As frmFacBancosPropios
Attribute frmMtoBancosPro.VB_VarHelpID = -1
Private WithEvents frmSocios As frmGesSocios
Attribute frmSocios.VB_VarHelpID = -1


Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1

'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Dim indCodigo As Integer 'indice para txtCodigo

Dim PrimeraVez As Boolean

Dim kCampo As Integer


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub chkImpAlbaran_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub chkImpAlbaran_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub




Private Sub chkSoloFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptarEstVentas_Click()
'Estadistica Ventas por meses
Dim campo As String
    
    InicializarVbles
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1

    
    'El campo AÑO es obligarotorio
    txtCodigo(53).Text = Trim(txtCodigo(53).Text)
    If txtCodigo(53).Text = "" Then
        MsgBox "Debe seleccionar una año para el informe.", vbInformation
        Exit Sub
    Else
        campo = "year({scafac.fecfactu})"
        cadFormula = campo & " = " & txtCodigo(53).Text
'        campo = campo & " = " & CInt(txtCodigo(53).Text) - 1
'        cadFormula = "(" & cadFormula & " OR " & campo & ")"
        
        'Parametro del año solicitado para el informe
        'Pasar el año solicitado como parametro
        cadParam = cadParam & "pAnyo=""" & "Año: " & txtCodigo(53).Text & """|"
        numParam = numParam + 1
    End If
    
    'Campo seleccion de un CLIENTE
    txtCodigo(8).Text = Trim(txtCodigo(8).Text)
    If txtCodigo(8).Text <> "" Then
        campo = "{scafac.codclien}"
        cadFormula = cadFormula & " AND (" & campo & " =" & txtCodigo(8).Text & ")"
        'Pasar el cliente solicitado como parametro
        cadParam = cadParam & "pDHCliente=""" & "Cliente: " & txtCodigo(8).Text & " - " & txtNombre(8).Text & """|"
    Else
        'Mostrar en el informe el total del Año Anterior
        campo = campo & " = " & CInt(txtCodigo(53).Text) - 1
        cadFormula = "(" & cadFormula & " OR " & campo & ")"
        
        cadParam = cadParam & "pDHCliente=""" & "Cliente: Todos" & """|"
    End If
    numParam = numParam + 1
    
    
    'Comprobar si hay registros para mostrar en el informe
    cadSelect = cadFormula
    If Not HayRegParaInforme("scafac", cadSelect) Then Exit Sub
    
    
    'Borro los datos temporales,por si acaso se hubiera quedado
    BorrarTempInformes
    
    'Generar la temporal con los totales por año, mes y cliente (tmpinformes)
    If Not TempVentasMeses(cadSelect, txtCodigo(53).Text) Then
        'Borrar los registros generados por el usuario de la temporal
        BorrarTempInformes
        Exit Sub
    End If
    
    cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
    Titulo = "Ventas por meses"
'    If Me.OptTipoInf(0).Value = True Then
        nomRPT = "rFacVentasxMesGra.rpt"
'    Else
'        Exit Sub
'        nomRPT = "rFacVentasxMesTex.rpt"
'    End If
    conSubRPT = False
    
    LlamarImprimir
    
    'Borrar los registros generados por el usuario de la temporal
    BorrarTempInformes
End Sub



Private Sub cmdAceptarFac_Click()
'Facturacion de Albaranes
Dim campo As String, cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean

    
    InicializarVbles
    cadFrom = ""
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtCodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCodigo(0).Text) = "" Then 'Banco propio
        MsgBox "El campo cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    
    
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    If OpcionListado <> 222 Then 'Facturas Ventas (FACTURACION)
                                 '222: Facturas de Mostrador/Rectificativa
        'Desde/Hasta Nº ALBARAN
        '-------------------------
        If txtCodigo(36).Text <> "" Or txtCodigo(37).Text <> "" Then
            campo = NomTabla & ".numalbar"
            cad = ""
            If Not PonerDesdeHasta(campo, "N", 36, 37, cad) Then Exit Sub
        End If
    
        'Desde/Hasta FECHA del ALBARAN
        '--------------------------------------------
        If txtCodigo(38).Text <> "" Or txtCodigo(39).Text <> "" Then
            'Para MySQL
            campo = "scaalb.fechaalb"
            cad = CadenaDesdeHastaBD(txtCodigo(38).Text, txtCodigo(39).Text, campo, "F")
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
            'Para Crystal Report
            campo = "{scaalb.fechaalb}"
            cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 38, 39, cad) Then Exit Sub
        End If
    
        'Cadena para seleccion D/H CLIENTE
        '----------------------------------------
        If txtCodigo(40).Text <> "" Or txtCodigo(41).Text <> "" Then
            campo = "scaalb.codclien"
            cad = ""
            If Not PonerDesdeHasta(campo, "N", 40, 41, cad) Then Exit Sub
        End If
    
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(42).Text <> "" Or txtCodigo(43).Text <> "" Then
            campo = "scaalb.codforpa"
            cad = " "
            If Not PonerDesdeHasta(campo, "N", 42, 43, cad) Then Exit Sub
        End If

    
        'Otros criterios de Seleccion
        '---------------------------------------------
        'Seleccionar de la Tabla de albaranes scaalb, solo los Albaranes que sean
        'del tipo:Ventas o Reparacion o Mantenimiento
    '    cad = " scaalb.codtipom='ALV' "
        cad = " scaalb.codtipom='" & CodClien & "' " 'filtrar por tipo de albaran segun llamado de Alb.Ventas o Alb. Reparacion
        'Solo lo añadimos a CadSelect porque vamos a Facturar y no a sacar un listado
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    
    
        'Seleccionar los Albanares de la Periodicidad indicada
        If txtCodigo(35).Text <> "" Then
            cad = " sclien.periodof=" & txtCodigo(35).Text
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
            cadFrom = " scaalb INNER JOIN sclien ON scaalb.codclien=sclien.codclien "
        End If
        
    Else
        'Facturar UNA solo
        If MsgBox("Generar la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        'en la llamada reutilizamos las vbles codclien y NumCod para guardar tipomov y numalbar.
        cadFormula = "{scaalb.codtipom}='" & CodClien & "' AND scaalb.numalbar=" & NumCod
        cadSelect = cadFormula
    End If
    
    
    
    cadSQL = cadSelect
                                                                
    'Pequeña comprobacion de los centros de coste
    If vEmpresa.TieneAnalitica Then
        cad = "select count(*) from slialb where codccost is null and (codtipom,numalbar) in ("
        cad = cad & "select codtipom,numalbar from scaalb where "
        cad = cad & cadSelect
        cad = cad & " AND  scaalb.factursn=1 )"
        cad = Replace(cad, "{", "(")
        cad = Replace(cad, "}", ")")
        NumRegElim = CInt(NumRegistros(cad))
        If NumRegElim > 0 Then
             cad = "Existen lineas de albaran(" & NumRegElim & ") sin asignar centro de coste"
             cad = cad & vbCrLf & vbCrLf & Space(30) & "¿Continuar?"
             If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If

    End If
    
                                                        
                                                                
                                                                
                                                                'Septiembre 2009
    'Seleccionar los Albaranes que tiene scaalb.factursn=1     y TENGAN lineas
    cad = " {scaalb.factursn=1} "
    
    'cad = cad & " and (scaalb.codtipom,scaalb.numalbar) in (select codtipom,numalbar from slialb group by codtipom,numalbar)"
    cad = cad & " and (scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb)"
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    AnyadirAFormula cadFormula, cad
    
        
    
    
    
    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    If cadFrom = "" Then cadFrom = " scaalb "
    cad = cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    If Not HayRegParaInforme(cadFrom, campo, True) Then
        MsgBox "Albaranes para facturar sin lineas", vbExclamation
        Exit Sub
    End If
    campo = ""
    'Verificar si con los criterios seleccionados (PARA VENTAS)
    'seleccionar si quedan en el rango Albaranes que no se van a Facturar
    'y mostrar mensaje
    If OpcionListado <> 222 Then
        'Seleccionar los Albaranes que tiene scaalb.factursn=0
        campo = " scaalb.factursn=0 "
        If Not AnyadirAFormula(cadSQL, campo) Then Exit Sub
        cadSQL = cad & " WHERE " & cadSQL
        If RegistrosAListar(cadSQL) > 0 Then
            'Mostrar los Albaranes que no se van a Facturar
            cadSQL = Replace(cadSQL, "count(*)", "scaalb.codtipom,scaalb.numalbar,scaalb.fechaalb,scaalb.codclien,scaalb.nomclien")
            frmMensajes.OpcionMensaje = 12
            frmMensajes.cadWHERE = cadSQL
            frmMensajes.Show vbModal
            If frmMensajes.vCampos = "0" Then Exit Sub
        End If
    End If
    
    cad = cad & " WHERE " & cadSelect
    

    
    
    'Pasar Albaranes a Facturas
    If InStr(cad, "sclien") <> 0 Then 'hay JOIN con sclien
        cad = Replace(cad, "count(*)", "scaalb.*, sclien.periodof")
    Else
        cad = Replace(cad, "count(*)", "*")
    End If







    'Albarananes EN B
    If CodClien = "ALZ" Then
        If Not AbrirConexionConta(True) Then
            cad = "Error MUY grave." & vbCrLf & "Error conectando con BD: " & vParamAplic.ContabilidadB
            MsgBox cad, vbCritical
            End
            Exit Sub
        End If
        CambiamosConta = True
    End If



    '--- Mostrar Barra de PRogreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador/Rectificativa
                                 '52: Facturas de Venta
                                 'Facturas Reparacion
        
        Me.Height = Me.Height + 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
        Me.FrameProgress.visible = True
        Me.FrameProgress.Top = 6250
        Me.ProgressBar1.Left = 200
        Me.ProgressBar1.Value = 0
        Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
        
        
        'Si vamos a facturar albaranes "B" tenemos que cerrar la conexion CONTA y abrirla contra la
        'Segunda BD que nos indica
        
    End If
    

    '--- Traspasa Albaranes a Facturas
    If OpcionListado = 222 And Me.EstaRecupFact = True Then
        '#### Laura: 14/11/2006 Recuperar facturas ALZIRA
        'comprobar q se ha introducido el nº de factura
        If Trim(txtCodigo(4).Text) = "" Then
            MsgBox "Debe introducir el nº de factura"
            Exit Sub
        End If
        'comprobar q la factura esta en un rango de recuperacion
        If Not (4255 <= CLng(txtCodigo(4).Text) And CLng(txtCodigo(4).Text) <= 5220) Then
            MsgBox "El Nº de factura no esta en el rango de recuperación."
            Exit Sub
        End If
        'comprobar q no exista ya ese nº de factura en aritaxi
        campo = "SELECT COUNT(*) FROM scafac WHERE "
        campo = campo & "codtipom='FAV' and numfactu=" & DBSet(txtCodigo(4).Text, "N") & " and year(fecfactu)=" & Year(txtCodigo(34).Text) '" and fecfactu=" & DBSet(txtCodigo(34).Text, "F")
        If Not (RegistrosAListar(campo) > 0) Then
            'comprobar si existe la factura en contabilidad
            campo = ""
            campo = ObtenerLetraSerie("FAV")
            If campo = "" Then Exit Sub
            campo = "SELECT COUNT(*) FROM cabfact WHERE numserie=" & DBSet(campo, "T")
            campo = campo & " AND codfaccl=" & txtCodigo(4).Text & " AND anofaccl=" & Year(txtCodigo(34).Text)
            
            If Not (RegistrosAListar(campo, conConta) > 0) Then
                'no existe en contabilidad recuperamos la factura y ya esta (no insertamos en tesoreria)
                TraspasoAlbaranesFacturas_RecuperaFac cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, txtCodigo(4).Text, Me.ProgressBar1, Me.lblProgess(1) 'Fecha de la factura, Cta Prevista de Cobro
            Else
                'si esiste
                'comprobar q el cliente es el mismo en la factura q vamos a recuperar
                'y en la factura de la conta
                If Not ComprobarCliente_RecuperarFac(cadSelect, txtCodigo(34).Text, txtCodigo(4).Text) Then Exit Sub
                'si existe en contabilidad recuperamos la factura y marcar como contabilizada
                TraspasoAlbaranesFacturas_RecuperaFac cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, txtCodigo(4).Text, Me.ProgressBar1, Me.lblProgess(1) 'Fecha de la factura, Cta Prevista de Cobro
                
                
            End If
        Else
            MsgBox "Ya existe la Factura en AriTaxi", vbExclamation
        End If
        '####################
    Else
        'proceso normal
         'Fecha de la factura, Cta Prevista de Cobro
         Screen.MousePointer = vbHourglass
         
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        campo = "Albaran: " & CodClien & " : " & NumCod
        LOG.Insertar 2, vUsu, campo
        Set LOG = Nothing
        '-----------------------------------------------------------------------------

        campo = txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
        'Para que no coja los albaranes sin lineas
        'cadSelect = cadSelect & " and (scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb )"
        If OpcionListado = 222 Then
            '[Monica]10/02/11: lo he quitado pq sino no me insertaba en tesoreria
            TraspasoAlbaranesFacturas cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo, False ' True
        Else
            TraspasoAlbaranesFacturas cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo, False
        End If
    End If
    Screen.MousePointer = vbDefault
    
    If CambiamosConta Then
       'Reestablecer la conexion con la antigua conta
       AbrirConexionConta False
    End If
    '--- Ocultar Barra de Progreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador
        Me.Height = Me.Height - 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
        Me.FrameProgress.visible = False
    Else
        'Cierro y salgo
        Unload Me
    End If
End Sub



'#### Laura 14/11/2006 Recuperar facturas ALZIRA
Private Function ComprobarCliente_RecuperarFac(cadSelAlb As String, FecFac As String, numFac As String) As Boolean
Dim Sql As String
Dim RS As ADODB.Recordset
Dim codMacta1 As String 'cliente factura aritaxi
Dim codMacta2 As String 'cliente factura conta
Dim LEtra As String

    On Error GoTo ErrCompCliente
    ComprobarCliente_RecuperarFac = False
    
    'codmacta del cliente del albaran a facturar en Aritaxi
    Sql = "select scaalb.codclien,sclien.codmacta"
    Sql = Sql & " from scaalb inner join sclien on scaalb.codclien=sclien.codclien "
    Sql = Sql & " Where " & cadSelAlb
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        codMacta1 = DBLet(RS!codmacta, "T")
    
    End If
    Set RS = Nothing
    
    
    'codmacta en la contabilidad
    LEtra = ObtenerLetraSerie("FAV")
    Sql = "SELECT codmacta FROM cabfact "
    Sql = Sql & " WHERE numserie=" & DBSet(LEtra, "T") & " AND codfaccl=" & numFac & " AND anofaccl=" & Year(FecFac)
    Set RS = New ADODB.Recordset
    RS.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        codMacta2 = DBLet(RS!codmacta, "T")
    End If
    Set RS = Nothing
    
    If codMacta1 <> "" And codMacta2 <> "" Then
        If codMacta1 = codMacta2 Then
            ComprobarCliente_RecuperarFac = True
        Else
            ComprobarCliente_RecuperarFac = False
            MsgBox "La cuenta contable en la factura de Contabilidad no coincide con la del cliente del Albaran", vbExclamation
        End If
    Else
        ComprobarCliente_RecuperarFac = False
        MsgBox "No se ha podido leer la cuenta contable del cliente", vbExclamation
    End If
    
    Exit Function
    
ErrCompCliente:
    ComprobarCliente_RecuperarFac = False
    MuestraError Err.Number, "Comprobar cliente", Err.Description
End Function
'#####


Private Sub cmdAceptarFacCli_Click()
'Facturacion de Albaranes
Dim campo As String, cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean

    
    InicializarVbles
    cadFrom = ""
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtCodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCodigo(0).Text) = "" Then 'Banco propio
        MsgBox "El campo cta. prevista de cobro debe tener valor.", vbExclamation
        Exit Sub
    End If
    
    
    
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    If OpcionListado <> 222 Then 'Facturas Ventas (FACTURACION)
                                 '222: Facturas de Mostrador/Rectificativa
        'Desde/Hasta Nº ALBARAN
        '-------------------------
        If txtCodigo(36).Text <> "" Or txtCodigo(37).Text <> "" Then
            campo = NomTabla & ".numalbar"
            cad = ""
            If Not PonerDesdeHasta(campo, "N", 36, 37, cad) Then Exit Sub
        End If
    
        'Desde/Hasta FECHA del ALBARAN
        '--------------------------------------------
        If txtCodigo(38).Text <> "" Or txtCodigo(39).Text <> "" Then
            'Para MySQL
            campo = "scaalbcli.fechaalb"
            cad = CadenaDesdeHastaBD(txtCodigo(38).Text, txtCodigo(39).Text, campo, "F")
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
            'Para Crystal Report
            campo = "{scaalbcli.fechaalb}"
            cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 38, 39, cad) Then Exit Sub
        End If
    
        'Cadena para seleccion D/H CLIENTE
        '----------------------------------------
        If txtCodigo(40).Text <> "" Or txtCodigo(41).Text <> "" Then
            campo = "scaalbcli.codclien"
            cad = ""
            If Not PonerDesdeHasta(campo, "N", 40, 41, cad) Then Exit Sub
        End If
    
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(42).Text <> "" Or txtCodigo(43).Text <> "" Then
            campo = "scaalbcli.codforpa"
            cad = " "
            If Not PonerDesdeHasta(campo, "N", 42, 43, cad) Then Exit Sub
        End If

    
        'Otros criterios de Seleccion
        '---------------------------------------------
        'Seleccionar de la Tabla de albaranes scaalb, solo los Albaranes que sean
        'del tipo:Ventas o Reparacion o Mantenimiento
    '    cad = " scaalb.codtipom='ALV' "
        cad = " scaalbcli.codtipom='" & CodClien & "' " 'filtrar por tipo de albaran segun llamado de Alb.Ventas o Alb. Reparacion
        'Solo lo añadimos a CadSelect porque vamos a Facturar y no a sacar un listado
        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    
    
        'Seleccionar los Albanares de la Periodicidad indicada
        If txtCodigo(35).Text <> "" Then
            cad = " scliente.periodof=" & txtCodigo(35).Text
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
            cadFrom = " scaalb INNER JOIN scliente ON scaalbcli.codclien=scliente.codclien "
        End If
        
    Else
        'Facturar UNA solo
        If MsgBox("Generar la factura?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
        'en la llamada reutilizamos las vbles codclien y NumCod para guardar tipomov y numalbar.
        cadFormula = "{scaalbcli.codtipom}='" & CodClien & "' AND scaalbcli.numalbar=" & NumCod
        cadSelect = cadFormula
    End If
    
    
    
    cadSQL = cadSelect
                                                                
    'Pequeña comprobacion de los centros de coste
    If vEmpresa.TieneAnalitica Then
        cad = "select count(*) from slialbcli where codccost is null and (codtipom,numalbar) in ("
        cad = cad & "select codtipom,numalbar from scaalbcli where "
        cad = cad & cadSelect
        cad = cad & " AND  scaalbcli.factursn=1 )"
        cad = Replace(cad, "{", "(")
        cad = Replace(cad, "}", ")")
        NumRegElim = CInt(NumRegistros(cad))
        If NumRegElim > 0 Then
             cad = "Existen lineas de albaran(" & NumRegElim & ") sin asignar centro de coste"
             cad = cad & vbCrLf & vbCrLf & Space(30) & "¿Continuar?"
             If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If

    End If
    
                                                        
                                                                
                                                                
                                                                'Septiembre 2009
    'Seleccionar los Albaranes que tiene scaalb.factursn=1     y TENGAN lineas
    cad = " {scaalbcli.factursn=1} "
    
    'cad = cad & " and (scaalb.codtipom,scaalb.numalbar) in (select codtipom,numalbar from slialb group by codtipom,numalbar)"
    cad = cad & " and (scaalbcli.codtipom,scaalbcli.numalbar) in (select distinct codtipom,numalbar from slialbcli)"
    If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
    AnyadirAFormula cadFormula, cad
    
        
    
    
    
    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    If cadFrom = "" Then cadFrom = " scaalbcli "
    cad = cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    If Not HayRegParaInforme(cadFrom, campo, True) Then
        MsgBox "Albaranes para facturar sin lineas", vbExclamation
        Exit Sub
    End If
    campo = ""
    'Verificar si con los criterios seleccionados (PARA VENTAS)
    'seleccionar si quedan en el rango Albaranes que no se van a Facturar
    'y mostrar mensaje
    If OpcionListado <> 222 Then
        'Seleccionar los Albaranes que tiene scaalb.factursn=0
        campo = " scaalbcli.factursn=0 "
        If Not AnyadirAFormula(cadSQL, campo) Then Exit Sub
        cadSQL = cad & " WHERE " & cadSQL
        If RegistrosAListar(cadSQL) > 0 Then
            'Mostrar los Albaranes que no se van a Facturar
            cadSQL = Replace(cadSQL, "count(*)", "scaalbcli.codtipom,scaalbcli.numalbar,scaalbcli.fechaalb,scaalbcli.codclien,scaalbcli.nomclien")
            frmMensajes.OpcionMensaje = 12
            frmMensajes.cadWHERE = cadSQL
            frmMensajes.Show vbModal
            If frmMensajes.vCampos = "0" Then Exit Sub
        End If
    End If
    
    cad = cad & " WHERE " & cadSelect
    

    
    
    'Pasar Albaranes a Facturas
    If InStr(cad, "scliente") <> 0 Then 'hay JOIN con sclien
        cad = Replace(cad, "count(*)", "scaalbcli.*, scliente.periodof")
    Else
        cad = Replace(cad, "count(*)", "*")
    End If







    'Albarananes EN B
    If CodClien = "ALZ" Then
        If Not AbrirConexionConta(True) Then
            cad = "Error MUY grave." & vbCrLf & "Error conectando con BD: " & vParamAplic.ContabilidadB
            MsgBox cad, vbCritical
            End
            Exit Sub
        End If
        CambiamosConta = True
    End If



    '--- Mostrar Barra de PRogreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador/Rectificativa
                                 '52: Facturas de Venta
                                 'Facturas Reparacion
        
        Me.Height = Me.Height + 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
        Me.FrameProgress.visible = True
        Me.FrameProgress.Top = 6250
        Me.ProgressBar1.Left = 200
        Me.ProgressBar1.Value = 0
        Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
        
        
        'Si vamos a facturar albaranes "B" tenemos que cerrar la conexion CONTA y abrirla contra la
        'Segunda BD que nos indica
        
    End If
    

    '--- Traspasa Albaranes a Facturas
    If OpcionListado = 222 And Me.EstaRecupFact = True Then
        '#### Laura: 14/11/2006 Recuperar facturas ALZIRA
        'comprobar q se ha introducido el nº de factura
        If Trim(txtCodigo(4).Text) = "" Then
            MsgBox "Debe introducir el nº de factura"
            Exit Sub
        End If
        'comprobar q la factura esta en un rango de recuperacion
        If Not (4255 <= CLng(txtCodigo(4).Text) And CLng(txtCodigo(4).Text) <= 5220) Then
            MsgBox "El Nº de factura no esta en el rango de recuperación."
            Exit Sub
        End If
        'comprobar q no exista ya ese nº de factura en aritaxi
        campo = "SELECT COUNT(*) FROM scafaccli WHERE "
        campo = campo & "codtipom='FAV' and numfactu=" & DBSet(txtCodigo(4).Text, "N") & " and year(fecfactu)=" & Year(txtCodigo(34).Text) '" and fecfactu=" & DBSet(txtCodigo(34).Text, "F")
        If Not (RegistrosAListar(campo) > 0) Then
            'comprobar si existe la factura en contabilidad
            campo = ""
            campo = ObtenerLetraSerie("FAV")
            If campo = "" Then Exit Sub
            campo = "SELECT COUNT(*) FROM cabfact WHERE numserie=" & DBSet(campo, "T")
            campo = campo & " AND codfaccl=" & txtCodigo(4).Text & " AND anofaccl=" & Year(txtCodigo(34).Text)
            
            If Not (RegistrosAListar(campo, conConta) > 0) Then
                'no existe en contabilidad recuperamos la factura y ya esta (no insertamos en tesoreria)
                TraspasoAlbaranesFacturas_RecuperaFac cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, txtCodigo(4).Text, Me.ProgressBar1, Me.lblProgess(1) 'Fecha de la factura, Cta Prevista de Cobro
            Else
                'si esiste
                'comprobar q el cliente es el mismo en la factura q vamos a recuperar
                'y en la factura de la conta
                If Not ComprobarCliente_RecuperarFac(cadSelect, txtCodigo(34).Text, txtCodigo(4).Text) Then Exit Sub
                'si existe en contabilidad recuperamos la factura y marcar como contabilizada
                TraspasoAlbaranesFacturas_RecuperaFac cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, txtCodigo(4).Text, Me.ProgressBar1, Me.lblProgess(1) 'Fecha de la factura, Cta Prevista de Cobro
                
                
            End If
        Else
            MsgBox "Ya existe la Factura en AriTaxi", vbExclamation
        End If
        '####################
    Else
        'proceso normal
         'Fecha de la factura, Cta Prevista de Cobro
         Screen.MousePointer = vbHourglass
         
        '------------------------------------------------------------------------------
        '  LOG de acciones.
        Set LOG = New cLOG
        campo = "Albaran Rectif.: " & CodClien & " : " & NumCod
        LOG.Insertar 2, vUsu, campo
        Set LOG = Nothing
        '-----------------------------------------------------------------------------

        campo = txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
        'Para que no coja los albaranes sin lineas
        'cadSelect = cadSelect & " and (scaalb.codtipom,scaalb.numalbar) in (select distinct codtipom,numalbar from slialb )"
        If OpcionListado = 222 Then
            '[Monica]10/02/11: lo he quitado pq sino no me insertaba en tesoreria
            TraspasoAlbaranesFacturasCli cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo, False ' True
        Else
            TraspasoAlbaranesFacturasCli cad, cadSelect, txtCodigo(34).Text, txtCodigo(0).Text, Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo, False
        End If
    End If
    Screen.MousePointer = vbDefault
    
    If CambiamosConta Then
       'Reestablecer la conexion con la antigua conta
       AbrirConexionConta False
    End If
    '--- Ocultar Barra de Progreso
    If OpcionListado <> 222 Then '222: Facturas Mostrador
        Me.Height = Me.Height - 300
        Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
        Me.FrameProgress.visible = False
    Else
        'Cierro y salgo
        Unload Me
    End If

End Sub

Private Sub cmdAceptarGenAlb_Click()
'Solicitar datos para Generar Albaran a partir de un Pedido
Dim cad As String

    'DAVID
    'Comprobar que me han puesto algun dato
    '-------------------------------------------------------------------
    cad = ""
    If txtCodigo(17).Text = "" Or txtCodigo(18).Text = "" Or txtCodigo(19).Text = "" Or txtCodigo(25).Text = "" Then cad = "M"
    If OpcionListado = 1000 Then
        If txtCodigo(5).Text = "" Then cad = "M"
        If txtNombre(5).Text = "" Then cad = "M"
    End If
    If txtNombre(17).Text = "" Or txtNombre(18).Text = "" Or txtNombre(19).Text = "" Then cad = "M"
    
    If cad <> "" Then
        MsgBox "Campos obligatorios ", vbExclamation
        Exit Sub
    End If
    
    
    
    cad = txtCodigo(17).Text & "|"
    cad = cad & txtCodigo(18).Text & "|"
    cad = cad & txtCodigo(19).Text & "|"
    cad = cad & txtCodigo(25).Text & "|"
    cad = cad & Me.chkImpAlbaran.Value & "|"
    cad = cad & Me.chkImpEtiq.Value & "|"
    cad = cad & Me.chkImpHojaExped.Value & "|"
    'mando el banco propio
    If OpcionListado = 1000 Then cad = cad & txtCodigo(5).Text & "|"
    
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub



Private Sub cmdAceptarPedxArtic_Click()
'41: Informe de Pedidos por Articulo
'44: Informe de Pedidos por Cliente
'49: Informe de Albaranes por Artículo
Dim campo As String
Dim cad As String
Dim Sql As String
Dim cadFormula2 As String
Dim cadSelect2 As String
Dim cadSelect3 As String
Dim indice As Integer


    InicializarVbles
    cadFormula2 = ""
    cadSelect2 = ""
    cadSelect3 = ""
    
    'Pasar nombre de la Empresa como parametro
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1


    '===================================================
    '================= FORMULA =========================
    
    'Cadena para seleccion Fechas de Pedido/Albaran/Factura
    '--------------------------------------------
    'Desde/Hasta FECHA
    'para el informe 227 fecha requerida
    If OpcionListado = 227 Or OpcionListado = 230 Then
        If txtCodigo(11).Text = "" Or txtCodigo(12).Text = "" Then
            MsgBox "Los campos D/H fecha factura deben tener valor.", vbInformation
            Exit Sub
        End If
        
        If DateDiff("d", txtCodigo(11).Text, txtCodigo(12).Text) > 365 Then
            MsgBox "El intervalo de fechas no puede ser superior a un año.", vbInformation
            Exit Sub
        End If
    End If
    
    If txtCodigo(11).Text <> "" Or txtCodigo(12).Text <> "" Then
        If OpcionListado = 227 Or OpcionListado = 228 Or OpcionListado = 230 Then
            campo = "{" & NomTabla & ".fecfactu}"
        ElseIf OpcionListado = 49 Then
            campo = "{" & NomTabla & ".fechaalb}"
        Else
            campo = "{" & NomTabla & ".fecpedcl}"
        End If
        cad = "pDHFecha=""Fecha: "
        If Not PonerDesdeHasta(campo, "F", 11, 12, cad) Then Exit Sub
        cadSelect = CadenaDesdeHastaBD(txtCodigo(11).Text, txtCodigo(12).Text, campo, "F")
        
        'Guardamos el periodo para calcular las ventas
        If OpcionListado = 227 Or OpcionListado = 230 Then
            cadFormula2 = cadFormula
            cadSelect2 = cadSelect
            'obtenemos el periodo anterior de ventas
            cad = "": Sql = ""
            If txtCodigo(11).Text <> "" Then cad = Day(txtCodigo(11).Text) & "/" & Month(txtCodigo(11).Text) & "/" & Year(txtCodigo(11).Text) - 1
            If txtCodigo(12).Text <> "" Then Sql = Day(txtCodigo(12).Text) & "/" & Month(txtCodigo(12).Text) & "/" & Year(txtCodigo(12).Text) - 1
            cadSelect3 = CadenaDesdeHastaBD(cad, Sql, campo, "F")
        
        ElseIf OpcionListado = 41 Or OpcionListado = 42 Then '42:Disponibilidad Stock
        'pasar D/H fecha como parametro para enlazar con la cabecera de pedidos proveedor
        'que esta como subinforme y que seleccione el mismo rango de fecha que
        'para la cabecera de pedidos de cliente
            If txtCodigo(11).Text <> "" Then
                cad = "pFechaD=" & "Date(" & Year(txtCodigo(11).Text) & ", " & Month(txtCodigo(11).Text) & ", " & Day(txtCodigo(11).Text) & ")"
            Else
                cad = "pFechaD=" & "Date(1900,01,01)"
            End If
            cadParam = cadParam & cad & "|"
            numParam = numParam + 1
            If txtCodigo(12).Text <> "" Then
                cad = "pFechaH=" & "Date(" & Year(txtCodigo(12).Text) & ", " & Month(txtCodigo(12).Text) & ", " & Day(txtCodigo(12).Text) & ")"
            Else
                cad = "pFechaH=" & "Date(9999,12,31)"
            End If
            cadParam = cadParam & cad & "|"
            numParam = numParam + 1
        End If
    End If
    
    'Cadena para seleccion ALMACEN
    '--------------------------------------------
    If Me.Frame9.visible Then
        If txtCodigo(13).Text <> "" Or txtCodigo(14).Text <> "" Then
            campo = "{" & NomTablaLin & ".codalmac}"
            'Parametro Desde/Hasta Almacen
            cad = "pDHAlmacen=""Almacen: "
            If Not PonerDesdeHasta(campo, "N", 13, 14, cad) Then Exit Sub
        End If
    End If
    
    
    'Cadena para seleccion ARTICULO
    '--------------------------------------------
    If Me.Frame8.visible Then
        If txtCodigo(15).Text <> "" Or txtCodigo(16).Text <> "" Then
            campo = "{" & NomTablaLin & ".codartic}"
            'Parametro Desde/Hasta Articulo
            cad = "pDHArticulo=""Artículo: "
             If Not PonerDesdeHasta(campo, "T", 15, 16, cad) Then Exit Sub
        End If
    End If
    
    'Cadena para seleccion SOCIO
    '--------------------------------------------
    If Me.Frame5.visible Then
        If txtCodigo(20).Text <> "" Or txtCodigo(21).Text <> "" Then
            campo = "{" & NomTabla & ".codclien}"
            'Parametro Desde/Hasta Cliente
            cad = "pDHCliente=""Socio: "
            If Not PonerDesdeHasta(campo, "N", 20, 21, cad) Then Exit Sub
        End If
    End If
    
    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If Me.Frame1.visible Then
        If txtCodigo(22).Text <> "" Or txtCodigo(23).Text <> "" Then
            campo = "{" & NomTabla & ".codclien}"
            'Parametro Desde/Hasta Cliente
            cad = "pDHCliente=""Cliente: "
            If Not PonerDesdeHasta(campo, "N", 22, 23, cad) Then Exit Sub
        End If
    End If
    
    'Cadena para seleccion TRABAJADOR
    '--------------------------------------------
    If Me.Frame12.visible Then
        If txtCodigo(2).Text <> "" Or txtCodigo(3).Text <> "" Then
            campo = "{scafac1.codtraba}"
            'Parametro Desde/Hasta Trabajador
            cad = "pDHTrabajador=""Trabajador: "
            If Not PonerDesdeHasta(campo, "N", 2, 3, cad) Then Exit Sub
        End If
    End If
    
    
    
    '227: Listado Ventas por cliente
    'Importe ventas superior a ....
    If Me.Frame10.visible Then
        cad = DBSet(txtCodigo(1).Text, "N")
        cadParam = cadParam & "pImporte=" & cad & "|"
        numParam = numParam + 1
            
        If txtCodigo(1).Text <> "" Then
            'seleccionar solo los clientes que el total de la BaseImp supere esa cantidad
            If cadSelect <> "" Then Sql = cadSelect2 & " AND "
            If OpcionListado = 230 Then
                cad = ObtenerClientes(cadSelect, cad, True)
            Else
                cad = ObtenerClientes(cadSelect, cad)
            End If
            cadSelect = Sql & cad
'            If cadSelect3 <> "" Then cadSelect3 = cadSelect3 & " AND "
'            cadSelect3 = cadSelect3 & cad
            If cadFormula2 <> "" Then cadFormula2 = cadFormula2 & " AND "
            cadFormula = cadFormula2 & cad
        End If
    End If
    
    
    If OpcionListado = 49 Then
        campo = ".numalbar"
'        cad = "{" & NomTabla & ".codtipom}='ALV'"
'        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
'        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        '-- Ahora en este informe hay mas posibilidades de selección [SERVICIOS]
        If vParamAplic.Servicios Then
            indice = cmbTipAlbaran(1).ListIndex
            If indice < 0 Then
                MsgBox "Debe seleccionar el tipo o los tipos de alabarán a procesar", vbExclamation
                Exit Sub
            Else
                Select Case indice
                    Case 0 ' solo ventas
                        cad = "{" & NomTabla & ".codtipom}='ALV'"
                        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Ventas)"
                    Case 1 ' solo servicios
                        cad = "{" & NomTabla & ".codtipom}='ALS'"
                        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Servicios)"
                    Case 2 ' ventas y servicios
                        cad = " ({" & NomTabla & ".codtipom}='ALV'" & _
                                " OR {" & NomTabla & ".codtipom}='ALS')"
                        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
                        Titulo = "Albaranes por artículo (Ventas y servicios)"
                End Select
            End If
        Else
            cad = "{" & NomTabla & ".codtipom}='ALV'"
            If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
            Titulo = "Albaranes por artículo (Ventas)"
        End If
        'Pasar nombre el título del informe
        cadParam = cadParam & "|pTitulo=""" & Titulo & """|"
        numParam = numParam + 1
    Else
        campo = ".numpedcl"
    End If

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If OpcionListado = 227 Then
        cad = NomTabla
        Titulo = "Ventas por Socio"
        nomRPT = "rFacVentasxClien.rpt"
        conSubRPT = False
    ElseIf OpcionListado = 230 Then
        cad = NomTabla
        Titulo = "Ventas por Cliente"
        nomRPT = "rFacVentasxCliente.rpt"
        conSubRPT = False
    ElseIf OpcionListado = 228 Then
        cad = NomTabla & " INNER JOIN scafac1 ON scafac.codtipom=scafac1.codtipom and scafac.fecfactu=scafac1.fecfactu and scafac.numfactu=scafac1.numfactu"
        Titulo = "Ventas por Trabajador"
        If Me.OptDetalle(2).Value = True Then 'Inf. Detalle
            nomRPT = "rFacVentasxTrabaDet.rpt"
            conSubRPT = True
        ElseIf Me.OptResumen.Value = True Then 'Inf. Resum
            nomRPT = "rFacVentasxTrabaRes.rpt"
            conSubRPT = False
        End If
    Else
       cad = NomTabla & " INNER JOIN " & NomTablaLin
       cad = cad & " ON " & NomTabla & campo & "=" & NomTablaLin & campo
       If OpcionListado = 49 Then _
       cad = cad & " AND " & NomTabla & ".codtipom=" & NomTablaLin & ".codtipom "
    End If
    
    If Not HayRegParaInforme(cad, cadSelect) Then Exit Sub
    
    
    If OpcionListado = 227 Then
        BorrarTempInformes
        
        'Pasar los datos a la tabla temporal tmpInformes y luego mostrar
        'el informe de esta tabla
        cadSelect2 = Replace(cadSelect2, "{", "")
        cadSelect2 = Replace(cadSelect2, "}", "")
        
        cadSelect3 = Replace(cadSelect3, "{", "")
        cadSelect3 = Replace(cadSelect3, "}", "")
        If Not TempVentasClientes(cadSelect, cadSelect2, cadSelect3) Then Exit Sub
        
        'Añadir como parametros el total del periodo que devuelve en cadSelect2
        'y añadir el parametro del total periodo anterior q devuelve en cadSelect3
        cadParam = cadParam & "pTotal=" & cadSelect2 & "|"
        numParam = numParam + 1
        cadParam = cadParam & "pTotalAnt=" & cadSelect3 & "|"
        numParam = numParam + 1
        
        
        'Añadir el parametro para el orden del informe
        'Orden del Informe
        If Me.OptOrdenCodclien.Value Then
            cad = "{tmpinformes.codigo1}"
            Sql = "Orden: Cod. cliente"
        ElseIf Me.OptOrdenNomclien.Value Then
            cad = "{tmpinformes.nombre1}"
            Sql = "Orden: Nombre cliente"
        ElseIf Me.OptOrdenVentas.Value Then
            cad = "{tmpinformes.importe5}"
            Sql = "Orden: Volumen ventas"
        End If
        cadParam = cadParam & "pOrden=" & cad & "|"
        numParam = numParam + 1
        cadParam = cadParam & "pCadOrden=""" & Sql & """|"
        numParam = numParam + 1
        
        
        'no le pasamos formula de seleccion porque los datos ya estan en la temporal
        'solo el usuario que genero la informacion en la temporal
        cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
        
    ElseIf OpcionListado = 230 Then
        BorrarTempInformes
        
        'Pasar los datos a la tabla temporal tmpInformes y luego mostrar
        'el informe de esta tabla
        cadSelect2 = Replace(cadSelect2, "{", "")
        cadSelect2 = Replace(cadSelect2, "}", "")
        
        cadSelect3 = Replace(cadSelect3, "{", "")
        cadSelect3 = Replace(cadSelect3, "}", "")
        If Not TempVentasClientes(cadSelect, cadSelect2, cadSelect3, True) Then Exit Sub
        
        'Añadir como parametros el total del periodo que devuelve en cadSelect2
        'y añadir el parametro del total periodo anterior q devuelve en cadSelect3
        cadParam = cadParam & "pTotal=" & cadSelect2 & "|"
        numParam = numParam + 1
        cadParam = cadParam & "pTotalAnt=" & cadSelect3 & "|"
        numParam = numParam + 1
        
        
        'Añadir el parametro para el orden del informe
        'Orden del Informe
        If Me.OptOrdenCodclien.Value Then
            cad = "{tmpinformes.codigo1}"
            Sql = "Orden: Cod. cliente"
        ElseIf Me.OptOrdenNomclien.Value Then
            cad = "{tmpinformes.nombre1}"
            Sql = "Orden: Nombre cliente"
        ElseIf Me.OptOrdenVentas.Value Then
            cad = "{tmpinformes.importe5}"
            Sql = "Orden: Volumen ventas"
        End If
        cadParam = cadParam & "pOrden=" & cad & "|"
        numParam = numParam + 1
        cadParam = cadParam & "pCadOrden=""" & Sql & """|"
        numParam = numParam + 1
        
        
        'no le pasamos formula de seleccion porque los datos ya estan en la temporal
        'solo el usuario que genero la informacion en la temporal
        cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
    
        
    ElseIf OpcionListado = 44 Then
        If Me.optOrdePed(0).Value Then
            cad = "{sliped.codartic}"
        Else
            cad = "{scaped.numpedcl}"
        End If
        cadParam = cadParam & "rOrden=" & cad & "|"
        numParam = numParam + 1
        
        'MArzo 2010
        Sql = ""
        If txtCodigo(6).Text <> "" Or txtCodigo(7).Text <> "" Then
            campo = "{scaped.codagent}"
            'Parametro Desde/Hasta agente
            cad = "@=""Agente: "
            If Not PonerDesdeHasta(campo, "N", 6, 7, cad) Then Exit Sub
            Sql = Mid(cad, 4)
        End If
        
        
        
    End If
    
    
    LlamarImprimir
End Sub


Private Sub cmdAceptarPreFac_Click()
'Prevision de Facturacion de Albaranes
Dim campo As String, cad As String
Dim b As Boolean
Dim indice As Integer

    InicializarVbles
    b = (OpcionListado = 50)
    
    'If (Not B) Or (B And codClien = "ALV") Then
        'Pasar nombre de la Empresa como parametro
        cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
        numParam = numParam + 1
    'End If
    
    
    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    If Trim(txtCodigo(26).Text) <> "" Or Trim(txtCodigo(27).Text) <> "" Then
        'If b And CodClien <> "ALV" Then
        'If b Then
        '    campo = "scaalb.fechaalb"
        '    cad = "FECHA: "
        '    cadFormula = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, campo, "F")
        '    cadParam = cadParam & AnyadirParametroDH(cad, 26, 27) & """|"
        'Else
            'Para MySQL
            campo = "scaalb.fechaalb"
            cadSelect = CadenaDesdeHastaBD(txtCodigo(26).Text, txtCodigo(27).Text, campo, "F")
            'Para Crystal Report
            campo = "{scaalb.fechaalb}"
            cad = "pDHFecha=""Fecha: "
            If Not PonerDesdeHasta(campo, "F", 26, 27, cad) Then Exit Sub
        'End If
    End If

    'Cadena para seleccion CLIENTE
    '--------------------------------------------
    If txtCodigo(28).Text <> "" Or txtCodigo(29).Text <> "" Then
        'If b Then
        '    campo = "scaalb.codclien"
        '    cad = "CLIENTE: "
        'Else
            campo = "{scaalb.codclien}"
            cad = "pDHCliente=""Cliente: "
        'End If
        If Not PonerDesdeHasta(campo, "N", 28, 29, cad) Then Exit Sub
    End If

    If b Then 'opcionlistado=50
        'Cadena para seleccion FORMA PAGO
        '--------------------------------------------
        If txtCodigo(30).Text <> "" Or txtCodigo(31).Text <> "" Then
           ' If b Then
           '     campo = "scaalb.codforpa"
           '     cad = "F. PAGO: "
           ' Else
                campo = "{scaalb.codforpa}"
                cad = "pDHForpa=""Forma Pago: "
           ' End If
            If Not PonerDesdeHasta(campo, "N", 30, 31, cad) Then Exit Sub
        End If
        
        'seleccionar los Albaranes de Venta/Repar/Mantenim.
        'seleccionamos tipo de movimiento segun albaran de venta o Reparacion (ALV,ALR)
        '-- Aqui es donde se seleccionaban los albaranes a mostrar en el informe, ahora
        '   como se pueden seleccionar diferentes combinaciones se modifica la carga de la
        '   selección (se queda en rem la antigua línea) [SERVICIOS]
        
        If vParamAplic.Servicios And CodClien <> "ALR" Then
            indice = cmbTipAlbaran(0).ListIndex
            If indice < 0 Then
                MsgBox "Debe seleccionar el tipo o los tipos de alabarán a procesar", vbExclamation
                Exit Sub
            Else
                Select Case indice
                    Case 0 ' solo ventas
                        cad = " {scaalb.codtipom}='ALV' "
                        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
                    Case 1 ' solo servicios
                        cad = " {scaalb.codtipom}='ALS' "
                        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
                    Case 2 ' ventas y servicios
                        cad = " ({scaalb.codtipom}='ALV'" & _
                                " OR {scaalb.codtipom}='ALS') "
                        If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
                        If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
                End Select
            End If
        Else
            cad = " {scaalb.codtipom}='" & CodClien & "' "
            If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        End If
        'Seleccionar los que esten marcados para facturar
        'Seleccionar solo aquellos que el campo scaalb.factursn=1
        If Me.chkSoloFacturar.Value = 1 Then
            cad = " {scaalb.factursn}=1 "
            If Not AnyadirAFormula(cadFormula, cad) Then Exit Sub
            If Not AnyadirAFormula(cadSelect, cad) Then Exit Sub
        End If
    Else
        'Cadena para seleccion AGENTE
        '--------------------------------------------
        If txtCodigo(32).Text <> "" Or txtCodigo(33).Text <> "" Then
            campo = "{scaalb.codagent}"
            cad = "pDHAgente="""
            If Not PonerDesdeHasta(campo, "N", 32, 33, cad) Then Exit Sub
        End If
        
        'Seleccionar solo aquellos que tienen Nº de Pedido para comparar los Plazos de Entrega
        campo = " NOT ISNULL({scaalb.numpedcl}) "
        If Not AnyadirAFormula(cadFormula, campo) Then Exit Sub
        If Not AnyadirAFormula(cadSelect, campo) Then Exit Sub
    End If
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme("scaalb", cadSelect) Then Exit Sub
    
    If OpcionListado = 51 Then
        Titulo = "Incumplimiento Plazos de Entrega"
        nomRPT = "rFacIncumPlazos.rpt"
        
    'ENERO 2009
    ElseIf OpcionListado = 50 Then
    'ElseIf OpcionListado = 50 And codClien = "ALV" Then
        If chkResumenForpa.Value = 1 Then
            'VAMOS A MOSTRAR LA HOJA RESUMEN DE FORMAS DE PAGO
            conn.Execute "DELETE FROM tmpinformes where codusu =" & vUsu.Codigo
        
            If Me.OptDetalle(0).Value Then
                Titulo = "SELECT scaalb.codforpa, sum(slialb.importel)," & vUsu.Codigo
                Titulo = Titulo & " FROM   ((scaalb scaalb INNER JOIN sclien sclien ON scaalb.codclien=sclien.codclien) INNER JOIN slialb slialb ON (scaalb.codtipom=slialb.codtipom) AND (scaalb.numalbar=slialb.numalbar)) INNER JOIN starif starif ON sclien.codtarif=starif.codlista"
            
            Else
                Titulo = "SELECT  codforpa ,sum(slialb.importel)," & vUsu.Codigo
                Titulo = Titulo & " FROM   slialb slialb INNER JOIN scaalb scaalb ON (slialb.codtipom=scaalb.codtipom) AND (slialb.numalbar=scaalb.numalbar)"
            End If
    
            If cadSelect <> "" Then Titulo = Titulo & " WHERE " & cadSelect
            
            Titulo = Titulo & " GROUP BY codforpa"
            Titulo = "INSERT INTO tmpinformes (codigo1,importe1,codusu) " & Titulo
            conn.Execute Titulo
        End If
    
    
        Titulo = "Previsión Facturación Ventas"
        If CodClien = "ALR" Then Titulo = Titulo & "(REP)"
        '-- Si estan activos los servicios hay diferentes posibilidades y el título
        '   las refleja, la variabele 'indice' lleva la información del combo seleccionado y
        '   ha sido cargada un poco más arriba [SERVICIOS]
        
        If vParamAplic.Servicios And CodClien <> "ALR" Then
            Select Case indice
                Case 0
                    Titulo = "Previsión Facturación Ventas"
                Case 1
                    Titulo = "Previsión Facturación Servicios"
                Case 2
                    Titulo = "Previsión Facturación Ventas y Servicios"
            End Select
        End If
        If Me.OptDetalle(3).Value Then Titulo = Titulo & "(Fact.)"
         

        conSubRPT = True
        If Me.OptDetalle(0).Value = True Then
            nomRPT = "rFacPrevFactDetalle.rpt"
        ElseIf Me.OptDetalle(1).Value = True Then
            nomRPT = "rFacPrevFactResum.rpt"
            
        Else
            'Nuevo Marzo 2009
            'Como se facturara, es decir, el primer nivel de agrupacion es el tipofact de scaalb
            nomRPT = "rFacPrevFactDetalleCole.rpt"
        End If
        
        cad = "pCodUsu=" & vUsu.Codigo & "|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        '-- Ahora el título depende de los tipos de albaranes seleccionados [SERVICIOS]
        cad = "pTitulo=""" & Titulo & """|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        
        '--  Mostrara , o no, el subreport con el resumen por forma pago
        cad = "pVerForpa=" & Abs(chkResumenForpa.Value) & "|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        '--- departamentos
        cad = "TieneDpto=" & Abs(vParamAplic.Departamento) & "|"
        cadParam = cadParam & cad
        numParam = numParam + 1
        
        
        
        
        On Error GoTo EPreFact
        cad = "delete from tmpstockfec where codusu=" & vUsu.Codigo
        conn.Execute cad
        
        'Insertar total bonificaciones por cliente,articulo en una TEMPORAL
        cad = "SELECT " & vUsu.Codigo & " as codusu,  slialb.codartic,scaalb.codclien,sum(slialb.cantidad) as stock "
        cad = cad & "FROM (((scaalb INNER JOIN slialb ON scaalb.codtipom=slialb.codtipom and scaalb.numalbar=slialb.numalbar) "
        cad = cad & " INNER JOIN sbonif ON slialb.codartic=sbonif.codartic ) "
        cad = cad & " INNER JOIN sclien ON scaalb.codclien=sclien.codclien) "
        cad = cad & " INNER JOIN starif ON sclien.codtarif=starif.codlista "
        cad = cad & "WHERE " & cadSelect
        cad = cad & " AND starif.bonifica=1 "
        cad = cad & " GROUP BY scaalb.codclien,slialb.codartic"
        
        cad = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock) " & cad
        conn.Execute cad
        
        b = False 'PARA QUE NO ENTRE EN LO DE ABAJO y vaya a imprimir
    End If
    
    'If b And codClien <> "ALV" Then 'OpcionListado = 50 'NO Imprime, mostrar resultado en pantalla
'    If b And CodClien <> "ALV" Then 'OpcionListado = 50 'NO Imprime, mostrar resultado en pantalla
'        frmMensajes.cadWhere = cadSelect
'        frmMensajes.vCampos = cadParam
'        frmMensajes.OpcionMensaje = 6 'Prefacturacion Albaranes
'        frmMensajes.Show vbModal
'    Else
        LlamarImprimir
   ' End If
    
    'If OpcionListado = 50 And CodClien = "ALV" Then
    If OpcionListado = 50 Then
        cad = "delete from tmpstockfec where codusu=" & vUsu.Codigo
        conn.Execute cad
    End If
EPreFact:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Informe Prefacturación", Err.Description
    End If
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
     
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 49, 227, 228, 230 '41: Informe de Pedidos por Articulo
                        '42: Informe de Disponibilidad de Stocks
                        '44: Informe de Pedidos por Cliente
                        '49: Informe de Albaranes por Articulo
                        '227: Inf. estadistica Ventas por socio
                        '230: Inf. estadistica Ventas por cliente
                PonerFoco txtCodigo(11)
            Case 43, 1000
                    '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                    '1000: Pedido a factura:  Piede ademas de los datos del albaran, la cta prevista
                PonerFoco txtCodigo(17)
            Case 50 '50: Prevision de Facturacion Albaranes (NO IMPRIME LISTADO)
                PonerFoco txtCodigo(26)
            Case 52, 222  '52: Facturacion de Albaranes
                         '222: Factura de Mostrador
                PonerFoco txtCodigo(34)
            Case 229 '229: Inf. estadistica ventas por meses
                PonerFoco txtCodigo(53)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim indFrame As Single

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me
    
    'buscar
    For kCampo = 0 To 26
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    For kCampo = 32 To 34
        Me.imgBuscarOfer(kCampo).Picture = frmPpal.imgIcoForms.ListImages(1).Picture
    Next kCampo
    
    'fecha
    Me.imgFecha(0).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    For kCampo = 6 To 7
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    For kCampo = 10 To 14
        Me.imgFecha(kCampo).Picture = frmPpal.imgIcoForms.ListImages(2).Picture
    Next kCampo
    
    
    
    

    'Ocultar todos los Frames de Formulario
    Me.FramePedxArtic.visible = False
    Me.FrameGenAlbaran.visible = False
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    Me.FrameEstVentas.visible = False
    
    CommitConexion
    
    NomTabla = "scaped"
    NomTablaLin = "sliped"
        
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
            
        Case 49, 227, 228, 230
                    '49: Informe de Albaranes por Articulo
                    '227: Inf. estadistica Ventas por cliente
                    '228: Inf. estadistica Ventas por trabjador
            PonerFramePedxArticVisible True, H, W
            indFrame = 2 'solo para el boton cancelar
            '-- Si está activada la opción de servicios, muestra los controles que permiten
            '   el tipo o tipos de albaranes a mostrar en el informe, en caso contrario los
            '   oculta para no liar [SERVICIOS]
            lblTipAlbaran(1).visible = False
            cmbTipAlbaran(1).visible = False
            
            If vParamAplic.Servicios And OpcionListado = 49 Then
                lblTipAlbaran(1).visible = True
                cmbTipAlbaran(1).visible = True
            End If
            
            If OpcionListado = 49 Then 'Albaranes de Venta
                NomTabla = "scaalb"
                NomTablaLin = "slialb"
            ElseIf OpcionListado = 227 Or OpcionListado = 228 Then
                NomTabla = "scafac"
                NomTablaLin = "slifac"
                
                'poner por defecto las fechas del ejercicio contable
                Me.txtCodigo(11).Text = vEmpresa.FechaIni
                Me.txtCodigo(12).Text = vEmpresa.FechaFin
            ElseIf OpcionListado = 230 Then
                NomTabla = "scafaccli"
                NomTablaLin = "slifaccli"
                
                'poner por defecto las fechas del ejercicio contable
                Me.txtCodigo(11).Text = vEmpresa.FechaIni
                Me.txtCodigo(12).Text = vEmpresa.FechaFin
                
                Me.txtCodigo(22).TabIndex = 2
                Me.txtCodigo(23).TabIndex = 3
                
                OptOrdenCodclien.Caption = "Cod. cliente"
                OptOrdenNomclien.Caption = "Nombre cliente"
                Me.Label4(18).Caption = "Mostrar Clientes con ventas superior a"
            End If
            
        Case 43, 1000
                '43: Generar Albaran desde Pedido (NO IMPRIME LISTADO)
                '1000:  Pedido a factura: pide la cta prevista de cobro
            
            W = 6515
            H = 5900
            PonerFrameVisible Me.FrameGenAlbaran, True, H, W
            txtCodigo(25).Text = Format(Now, "dd/mm/yyyy")
            indFrame = 3
            chkImpAlbaran.Caption = "Impimir "
            If OpcionListado = 1000 Then
                Label4(32).Caption = "Fec. FACTURA"
                Label3.Caption = "FACTURAR pedido"
                chkImpAlbaran.Caption = chkImpAlbaran.Caption & "FACTURA"
            Else
                Label4(32).Caption = "Fecha albarán"
                chkImpAlbaran.Caption = chkImpAlbaran.Caption & "albaran"
                If NumCod = "REP" Then
                    Label3.Caption = "Pasar Reparación a Albaran"
                Else
                    Label3.Caption = "Pasar Pedido a Albaran"
                End If
            End If
            FramepedidoFactura.visible = (OpcionListado = 1000)
            
            '- Ver si hay articulo portes para imprimir hoja Expedicion
            If vParamAplic.ArtPortes <> "" Then
                Me.chkImpHojaExped.Value = 1
            Else
                Me.chkImpHojaExped.Value = 0
            End If
        
        
        Case 50 '50: Prevision Facturacion de Albaranes (NO IMPRIME LISTADO)
            PonerFramePreFacVisible True, H, W
            indFrame = 5 'solo para el boton cancelar
            '-- Si está activada la opción de servicios, muestra los controles que permiten
            '   el tipo o tipos de albaranes a mostrar en el informe, en caso contrario los
            '   oculta para no liar [SERVICIOS]
            lblTipAlbaran(0).visible = False
            cmbTipAlbaran(0).visible = False
            If vParamAplic.Servicios Then
                lblTipAlbaran(0).visible = CodClien <> "ALR"
                cmbTipAlbaran(0).visible = CodClien <> "ALR"
                lblTipAlbaran(0).Top = cmdAceptarPreFac.Top
                cmbTipAlbaran(0).Top = cmdAceptarPreFac.Top
            End If
            chkResumenForpa.visible = OpcionListado = 50
        Case 52, 222
                    '52: Facturacion de Albaranes
                    '222: Factura de Mostrador
                    
            PonerFrameFacVisible True, H, W
            txtCodigo(34).Text = Format(Now, "dd/mm/yyyy")
            txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            indFrame = 6
            
            NomTabla = "scaalb"
            NomTablaLin = "slialb"
            
            'Si es facturacion directa: 222 oculto el frame y muestro el albaran que voy a facturar
            Frame4.visible = (OpcionListado = 52)
            If OpcionListado = 52 Then
                Label10(10).Caption = ""
                Me.Frame15.Top = 5040
            Else
                Label10(10).Caption = "Albarán:     " & CodClien & "   " & NumCod
                Me.Frame15.Top = 1800
            End If
            
            
            If Escliente And OpcionListado = 222 Then
                cmdAceptarFacCli.visible = True
                cmdAceptarFacCli.Enabled = True
                cmdAceptarFac.visible = False
                cmdAceptarFac.Enabled = False
            End If
            
        Case 229 '229: Inf. estadistica ventas por mes
            H = 4000
            W = 7035
            PonerFrameVisible Me.FrameEstVentas, True, H, W
            indFrame = 8
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub



Private Sub frmB_Selecionado(CadenaDevuelta As String)
    cadFormula = CadenaDevuelta
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmMtoAgente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Agente
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoAlmacen_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Almacen
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoArticulo_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoBancosPro_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Bancos Propios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmMtoCliente_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFEnvio_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Envio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoFPago_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pabo
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMtoTraba_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Trabajadores
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSocios_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de socios
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
            
        Case 20, 21, 27, 28 'Cod. CLIENTE
            Select Case Index
                Case 14, 15: indCodigo = Index + 14
                Case 20, 21: indCodigo = Index + 20
                Case 27, 28: indCodigo = Index + 21
            End Select
            Set frmMtoCliente = New frmFacClientes
            frmMtoCliente.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoCliente.Show vbModal
            Set frmMtoCliente = Nothing
            
        Case 25, 26
            Select Case Index
                Case 25: indCodigo = 22
                Case 26: indCodigo = 23
            End Select
            Set frmMtoCliente = New frmFacClientes
            frmMtoCliente.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoCliente.Show vbModal
            Set frmMtoCliente = Nothing
        
        Case 14, 15, 11, 12, 32 'socios
            Select Case Index
                Case 14, 15: indCodigo = Index + 14
                Case 11, 12: indCodigo = Index + 9
                Case 32: indCodigo = 8
            End Select
            Set frmSocios = New frmGesSocios
            frmSocios.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmSocios.Show vbModal
            Set frmSocios = Nothing
        Case 4, 5 'Cod. ALMACEN
            If Index = 4 Then indCodigo = 13
            If Index = 5 Then indCodigo = 14
            Set frmMtoAlmacen = New frmAlmAlPropios
            frmMtoAlmacen.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoAlmacen.Show vbModal
            Set frmMtoAlmacen = Nothing
            
        Case 6, 7 'Cod. ARTICULO
            If Index = 6 Then
                indCodigo = 15
            Else
                indCodigo = 16
            End If
            Set frmMtoArticulo = New frmAlmArticulos
            frmMtoArticulo.DatosADevolverBusqueda2 = "@1@"
            frmMtoArticulo.Show vbModal
            Set frmMtoArticulo = Nothing
        
        Case 1, 2, 8, 9 'cod. TRABAJADOR
            Select Case Index
                Case 1, 2: indCodigo = Index + 1
                Case 8, 9: indCodigo = Index + 9
            End Select
            
            Set frmMtoTraba = New frmAdmTrabajadores
            frmMtoTraba.DatosADevolverBusqueda = "0|1|"
            frmMtoTraba.Show vbModal
            Set frmMtoTraba = Nothing
            
        Case 10 'Cod. Forma de Envio
            indCodigo = 19
            Set frmMtoFEnvio = New frmFacFormasEnvio
            frmMtoFEnvio.DatosADevolverBusqueda = "0|1|"
            frmMtoFEnvio.Show vbModal
            Set frmMtoFEnvio = Nothing
            
        Case 16, 17, 22, 23, 29, 30 'Forma de PAGO
            Select Case Index
                Case 16, 17: indCodigo = Index + 14
                Case 22, 23: indCodigo = Index + 20
                Case 29, 30: indCodigo = Index + 21
            End Select
            Set frmMtoFPago = New frmFacFormasPago
            frmMtoFPago.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoFPago.Show vbModal
            Set frmMtoFPago = Nothing
            
        Case 3, 13, 18, 19 'cod AGENTE
            If Index <= 13 Then
                'D/H agente para pedido x cliente
                'MARZO 2010
                indCodigo = 7
                If Index = 3 Then indCodigo = 6
            Else
                indCodigo = Index + 14
            End If
            Set frmMtoAgente = New frmFacAgentesCom
            frmMtoAgente.DatosADevolverBusqueda = "0|1|"
            If Not IsNumeric(txtCodigo(indCodigo).Text) Then txtCodigo(indCodigo).Text = ""
            frmMtoAgente.Show vbModal
            Set frmMtoAgente = Nothing
            
        Case 0, 24, 31 'Bancos Propios
            indCodigo = 0
            If Index = 31 Then
                indCodigo = 52
            ElseIf Index = 0 Then indCodigo = 5
            End If
            Set frmMtoBancosPro = New frmFacBancosPropios
            frmMtoBancosPro.DatosADevolverBusqueda = "0|1|"
            frmMtoBancosPro.Show vbModal
            Set frmMtoBancosPro = Nothing
        
        Case 54
            indCodigo = 54
            AbrirBuscaGrid indCodigo
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Set frmF = New frmCal
   frmF.Fecha = Now
   
   Select Case Index
        Case 0 'Frame Pasar Pedido -> Albaran
            indCodigo = 25
        Case 1 'framePedidos
            indCodigo = 3 'Desde
        Case 2 'framePedidos
            indCodigo = 4 'Hasta
        
        Case 6 'FramePedxArtic
            indCodigo = 11 'Fecha Desde
        Case 7 'FramePedxArtic
            indCodigo = 12 'Fecha Hasta
        Case 9 'FramePedCompras
            indCodigo = 24 'Fecha Hasta
        Case 10 'FramePreFacturar
            indCodigo = 26
        Case 11 'FramePreFacturar
            indCodigo = 27
        Case 12 'Frame Factura
            indCodigo = 38
        Case 13 'Frame Factura
            indCodigo = 39
        Case 14 'FrameFactura
            indCodigo = 34
   End Select
   
   PonerFormatoFecha txtCodigo(indCodigo)
   If txtCodigo(indCodigo).Text <> "" Then frmF.Fecha = CDate(txtCodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtCodigo(indCodigo)
End Sub







Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codCampo As String, NomCampo As String
Dim Tabla As String
      
    Select Case Index
        Case 1 'Importe (Decimal(12,2))
            PonerFormatoDecimal txtCodigo(Index), 1
            
        Case 0, 5, 52 'Bancos Propios
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sbanpr", "nombanpr", "codbanpr", "Bancos Propios")
                If txtCodigo(Index).Text <> "" And txtNombre(Index).Text <> "" Then
                    txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
                Else
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
        
        'FECHA Desde Hasta
        Case 11, 12, 25, 26, 27, 34, 38, 39, 44
            If txtCodigo(Index).Text <> "" Then
                PonerFormatoFecha txtCodigo(Index)
                If Index = 34 And txtCodigo(Index).Text <> "" Then _
                    txtCodigo(39).Text = Format(CDate(txtCodigo(34).Text) - 1, "dd/mm/yyyy")
            End If
           
'            'Fecha entrega para Pedido. Poner la semana
'            If Index = 26 Then txtNombre(4).Text = CalculaSemana(CDate(txtCodigo(26).Text))
        
        Case 53 'AÑO
             PonerFormatoEntero txtCodigo(Index)
        
        Case 36, 37  'Nº de Pedido / Albaran
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000000")
            End If
            

        Case 35 'Periodicidad Facturacion
            PonerFormatoEntero txtCodigo(Index)

        Case 8, 20, 21, 28, 29, 40, 41, 48, 49 'Cod. Socio
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomclien"
                Tabla = "sclien"
                codCampo = "codclien"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, "Cliente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 22, 23 ' codigo de cliente
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomclien"
                Tabla = "scliente"
                codCampo = "codclien"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, "Cliente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 13, 14 'ALMACEN
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomalmac"
                Tabla = "salmpr"
                codCampo = "codalmac"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, "Almacen")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
     
        Case 2, 3, 17, 18, 47 'Cod. Trabajador
            If PonerFormatoEntero(txtCodigo(Index)) Then
                NomCampo = "nomtraba"
                Tabla = "straba"
                codCampo = "codtraba"
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, "Trabajador")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 19 'Cod. Envio
            NomCampo = "nomenvio"
            Tabla = "senvio"
            codCampo = "codenvio"
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, Tabla, NomCampo, codCampo, "Forma de Envío")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
            
        Case 30, 31, 42, 43, 50, 51 'Cod. Formas de PAGO
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sforpa", "nomforpa", "codforpa", "Formas de Pago")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
        Case 6, 7, 32, 33 'AGENTE
            If PonerFormatoEntero(txtCodigo(Index)) Then
                txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sagent", "nomagent", "codagent", "Agente")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "0000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 45 'TIPO CONTRATO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "stipco", "nomtipco", "codtipco", "Tipo Contrato", "T")
            
        Case 46 'MES a facturar
            If PonerFormatoEntero(txtCodigo(Index)) Then
                'Comprobar que el mes es correcto, valores entre 1-12
                devuelve = txtCodigo(Index).Text
                If (CByte(devuelve) >= 1) And (CByte(devuelve) <= 12) Then
                    txtNombre(Index).Text = UCase(MonthName(CLng(devuelve)))
                Else
                    MsgBox "El valor introducido no es un MES válido.(1-12).", vbInformation
                    txtCodigo(Index).Text = ""
                    PonerFoco txtCodigo(Index)
                End If
            Else
                txtNombre(Index).Text = ""
            End If
            
            
        Case 54
            'Centro de coste
            txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
            codCampo = ""
            If txtCodigo(Index).Text <> "" Then
                
                codCampo = "nomccost"
                Tabla = DevuelveDesdeBD(conConta, "codccost", "cabccost", "codccost", txtCodigo(Index).Text, "T", codCampo)
            
                If Tabla = "" Then
                    MsgBox "No existe el centro de coste: " & txtCodigo(Index).Text, vbExclamation
                    
                End If
                If codCampo = "nomccost" Then codCampo = ""
                txtCodigo(Index).Text = Tabla
            End If
            txtNombre(Index).Text = codCampo
            
            
        '##### Recuperar facturas ALZIRA
        Case 4 'nº factura
            PonerFocoBtn Me.cmdAceptarFac
        '#####
    End Select
End Sub



Private Sub PonerFramePedxArticVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Informe Pedidos por Articulo Visible y Ajustado al Formulario, y visualiza los controles
'necesario para cada Informe de Ofertas

    'MARZO 2010
    'Los botones de aceptar cancelar
    H = 4800
    If OpcionListado = 44 Then H = 7080
    cmdAceptarPedxArtic.Top = H
    Me.cmdCancel(2).Top = H


    H = 5415
    'Marzo 2010
    If OpcionListado = 44 Then H = 7575
    W = 7515
    

    
        
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePedxArtic, visible, H, W
    
    If visible = True Then
        Me.Frame5.visible = (OpcionListado = 44) Or (OpcionListado = 227)  'D/H socio
        Me.Frame1.visible = (OpcionListado = 230) 'D/H cliente
        'D/H Artículo
        Me.Frame8.visible = (OpcionListado <> 44) And (OpcionListado <> 227) And (OpcionListado <> 228) And (OpcionListado <> 230)
        Me.Frame9.visible = (OpcionListado <> 227 And OpcionListado <> 228 And OpcionListado <> 230) 'D/H Almacen
        Me.Frame10.visible = (OpcionListado = 227) Or (OpcionListado = 230)
        Me.Frame12.visible = (OpcionListado = 228)
        FrameOrden2.visible = (OpcionListado = 44)
        FramepedxClien.visible = (OpcionListado = 44)
        'Para que salga
        
        If OpcionListado = 44 Then 'Informe Pedido por cliente
            Me.Frame5.Top = 3120
            Me.Frame5.Left = 500
            Me.Label1.Caption = "Pedidos por Cliente"
            '
            FramepedxClien.Top = 4440
            FramepedxClien.Left = 500
        ElseIf OpcionListado = 227 Then 'Inf. Estadistica ventas x cliente
            Me.Frame5.Top = 1800
            Me.Frame5.Left = 390 '500
            Me.Frame10.Top = 2800
            Me.Label1.Caption = "Ventas por Socio"
            Label4(4).Caption = "Fecha Factura"
            Me.cmdAceptarPedxArtic.Top = 4650
            Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        ElseIf OpcionListado = 228 Then 'Inf. Estadistica ventas x trabajador
            Me.Frame12.Top = 1900
            Me.Frame12.Left = 500
            Me.Label1.Caption = "Ventas por Trabajador"
            Label4(4).Caption = "Fecha Factura"
            Me.cmdAceptarPedxArtic.Top = 4150
            Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        ElseIf OpcionListado = 230 Then
            Me.Frame1.Top = 1800
            Me.Frame1.Left = 390 '500
            Me.Frame10.Top = 2800
            Me.Label1.Caption = "Ventas por Cliente"
            Label4(4).Caption = "Fecha Factura"
            Me.cmdAceptarPedxArtic.Top = 4650
            Me.cmdCancel(2).Top = Me.cmdAceptarPedxArtic.Top
        Else
            Me.Frame8.Top = 3120
            Me.Frame8.Left = 500
            If OpcionListado = 41 Then
                Me.Label1.Caption = "Pedidos por Artículo"
            ElseIf OpcionListado = 42 Then
                Me.Label1.Caption = "Disponibilidad de Stocks"
            ElseIf OpcionListado = 49 Then
                Me.Label1.Caption = "Albaranes por Artículo"
                Label4(4).Caption = "Fecha Albaran"
                Me.Frame8.Top = 3000
                Me.Frame8.Left = 400
            End If
        End If
    End If
End Sub


Private Sub PonerFramePreFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim b As Boolean
Dim cad As String

    H = 5600
    If OpcionListado = 51 Then 'Inf. Incum. plazos entrega
        H = 5300
        Me.cmdAceptarPreFac.Top = 4600
        Me.cmdCancel(5).Top = Me.cmdAceptarPreFac.Top
    End If
    W = 7040
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePreFacturar, visible, H, W
    If visible = True Then
        b = (OpcionListado = 50)
        Label4(41).visible = b
        Me.imgBuscarOfer(16).visible = b
        Me.imgBuscarOfer(17).visible = b
        Me.txtCodigo(30).visible = b
        Me.txtCodigo(31).visible = b
        Me.txtNombre(30).visible = b
        Me.txtNombre(31).visible = b
        Me.Frame6.visible = Not b
        Me.Frame6.Top = 2900
        Me.Frame6.Left = 460
        'solo albaranes a facturar
        Me.chkSoloFacturar.visible = b
        Me.chkSoloFacturar.Value = 1
        
        'Detalle o resumen
        Me.Frame7.visible = b And CodClien = "ALV"
        Me.Frame7.visible = b 'And CodClien = "ALV"
        Me.OptDetalle(0).Value = True
        
        If Not b Then
            Me.Label9(0).Caption = "Incum. plazos entrega"
        Else 'Prevision Facturacion
            Select Case CodClien 'aqui guardamos el tipo de movimiento
                Case "ALV": cad = "" ' antes "(Ventas)" [SERVICIOS]
                Case "ALR": cad = "(Reparaciones)"
                Case "ALM": cad = "(Mantenimientos)"
            End Select
            Me.Label9(0).Caption = "Previsión de facturación " & cad
        End If
    End If
End Sub


Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim cad As String

    H = 7100 + 180
    W = 7480
    
    If visible = True Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "ALV": cad = "(Ventas)"
            Case "ALR": cad = "(Reparaciones)"
            Case "ALM", "ART":
                If CodClien = "ALM" Then
                    cad = "(Mostrador)"
                Else
                    cad = "(Rectificativa)"
                End If
                'Me.Frame3.Top = 1200
                Me.Frame4.visible = False
                H = 4000
                Me.cmdAceptarFac.Top = 3200
                Me.cmdCancel(6).Top = Me.cmdAceptarFac.Top
            Case "ALS": cad = "(Servicios)"
                
                
        End Select
        '#### Laura Recuperar facturas ALZIRA
        'nº de factura solo visible si estamos recuperando facturas
        Me.Label10(9).visible = Me.EstaRecupFact And OpcionListado = 222
        Me.txtCodigo(4).visible = Me.EstaRecupFact And OpcionListado = 222
        If Me.EstaRecupFact And OpcionListado = 222 Then txtCodigo(0).Text = "001"
        
        Me.Label10(0).Caption = "Facturación de Albaranes " & cad
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
End Sub

Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
On Error Resume Next

    If txtCodigo(indD).Text <> "" And txtCodigo(indH).Text <> "" Then
        If txtCodigo(indD).Text = txtCodigo(indH).Text Then
            cad = cad & txtCodigo(indD).Text
            If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
            AnyadirParametroDH = cad
            Exit Function
        End If
    End If
    
    If txtCodigo(indD).Text <> "" Then
        cad = cad & "desde " & txtCodigo(indD).Text
        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
    End If
    If txtCodigo(indH).Text <> "" Then
        cad = cad & "  hasta " & txtCodigo(indH).Text
        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
    End If
    AnyadirParametroDH = cad
End Function


Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
Dim devuelve As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If Tipo <> "F" Then
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
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


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
           Case 15, 16 'ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), conAri, "sartic", "nomartic", "codartic", "Articulo", "T")
            If txtNombre(Index).Text = "" And txtCodigo(Index) <> "" Then Cancel = True
    End Select
End Sub




Private Function ObtenerClientes(cadW As String, Importe As String, Optional Escliente As Boolean) As String
Dim Sql As String
Dim RS As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    Sql = "select codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    If Escliente Then
        Sql = Sql & " From scafaccli "
    Else
        Sql = Sql & " From scafac "
    End If
    If cadW <> "" Then Sql = Sql & " where " & cadW
    Sql = Sql & " group by codclien "
    If Importe <> "" Then Sql = Sql & "having baseimp>" & Importe
    
    Set RS = New ADODB.Recordset
    RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not RS.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            Sql = Sql & RS!CodClien & ","
'        End If
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    If Sql <> "" Then
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        If Escliente Then
            Sql = "( {scafaccli.codclien} IN [" & Sql & "] )"
        Else
            Sql = "( {scafac.codclien} IN [" & Sql & "] )"
        End If
    End If
    ObtenerClientes = Sql
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function



Private Sub txtCSB_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub AbrirBuscaGrid(Op As Integer)
    
    Set frmB = New frmBuscaGrid
    cadFormula = "" 'Aqui metera el valor
    If Op = 54 Then
        'CEntro de coste
        
        frmB.vCampos = "Codigo|cabccost|codccost|T||20·Descripción|cabccost|nomccost|T||70·"
        frmB.vTabla = "cabccost"
        frmB.vSQL = ""
       
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = "Centros de coste"
        frmB.vselElem = 0
        frmB.vConexionGrid = conConta
    
    End If
    frmB.Show vbModal
    Set frmB = Nothing
    
    
    If cadFormula <> "" Then
        'Ha devuelto algun dato
        If Op = 54 Then
            txtCodigo(Op).Text = RecuperaValor(cadFormula, 1)
            txtNombre(Op).Text = RecuperaValor(cadFormula, 2)
        End If
    End If
End Sub
