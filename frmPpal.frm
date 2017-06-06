VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmPpal 
   BackColor       =   &H00858585&
   Caption         =   "Aritaxi"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   12780
   Icon            =   "frmPpal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageListB 
      Left            =   4920
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8688
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":909A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":AED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B8E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C2F4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListPpal 
      Left            =   360
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CD06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DD98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EE2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":FEBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":10F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":129D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":13A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":14AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":15B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":16C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":17CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":18D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":19DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1AE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1BEF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1CF84
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1E016
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":1F0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2013A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":21ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2832E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2C830
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":2D242
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":30634
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":36E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3D6F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3E78A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":3F81C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":408AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":41940
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":429D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":49234
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":4A2C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":50B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":51BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":52C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":53CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":54D70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   40
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Artículos"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Movimientos Art."
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clientes"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Proveedores"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Socios"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Histórico de Llamadas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Clientes"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaranes Socios"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Socios"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedidos Proveedor"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Albaran Proveedor"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Factura Proveedor"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recepción Facturas Prov."
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Liquidación Socio"
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Publicidad Socio"
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Cuotas Socio"
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Clientes"
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Facturas Publicidad Clientes"
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button32 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nº Serie"
         EndProperty
         BeginProperty Button33 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button34 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button35 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button36 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button37 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar empresa"
         EndProperty
         BeginProperty Button38 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button39 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button40 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   1
      Top             =   8430
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpal.frx":55E02
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14473
            Text            =   "asdasd"
            TextSave        =   "asdasd"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "10:20"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   5640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":593C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":5B0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":61374
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":61D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":62798
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":64F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":65824
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":660FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":669D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":672B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":67CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6811E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":68230
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":68342
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":68454
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6876E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6E390
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6EDA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6F7B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":6F8C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":702D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":70CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":716FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":71A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":71D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":72182
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":725D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":72A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":72E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":732CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7371C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":73A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":73B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":73EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":741C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":74A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":75378
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":75692
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":757EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":75B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":76518
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":76F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7793C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7834E
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":78D60
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListTPV 
      Left            =   360
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":79772
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7B104
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7CA96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7E428
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":7FDBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":830DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":84A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":8B2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":90AC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMAIL 
      Left            =   420
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":97326
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":97778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":97BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9801C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9846E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":988C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":98D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":99164
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":995B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":9F850
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A0262
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":A64FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ACD5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B35C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":B9E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C0684
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":C6EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CD748
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CDB9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CDFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CE43E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CE890
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CECE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":CF134
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D4D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D5BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D5EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D61DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D64F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   6240
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   6240
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   6240
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgListComun32 
      Left            =   7170
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   8640
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   8640
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun1 
      Left            =   7080
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D6810
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":D851A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DE7C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DF1D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":DFBE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E2396
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E2C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E354A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E3E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E46FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E5110
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E556A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E567C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E578E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E58A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":E5BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EB7DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EC1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ECC00
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ECD12
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":ED724
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EE136
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EEB48
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EEE62
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EF17C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EF5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EFA20
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":EFE72
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F02C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F0716
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F0E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F0FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F12F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F1610
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F27C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F2ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F2C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F2F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F3964
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4376
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F4D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F579A
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F61AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcoForms 
      Left            =   6300
      Top             =   5490
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F6BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F75D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F766B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal.frx":F807D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnConfiguracion 
      Caption         =   "C&onfiguración"
      Begin VB.Menu mnConfParamGenerales 
         Caption         =   "Datos &Empresa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnConfParamAplic 
         Caption         =   "Parámetros &Aplicación"
      End
      Begin VB.Menu mnConTMovimiento 
         Caption         =   "Tipos &Movimiento"
      End
      Begin VB.Menu mnConfParamRpt 
         Caption         =   "Tipos de &Documentos"
      End
      Begin VB.Menu mnAridoc1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnConfManteUsuarios 
         Caption         =   "Mantenimiento &Usuarios"
         HelpContextID   =   2
      End
      Begin VB.Menu mnNuevaEmpresa 
         Caption         =   "Creacion &nueva empresa"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnUsuarios 
         Caption         =   "Nuevo U&suario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnPedirPwd 
         Caption         =   "Password requerido"
         Visible         =   0   'False
      End
      Begin VB.Menu mnCambioEmpresa 
         Caption         =   "Cambiar Em&presa"
         HelpContextID   =   2
      End
      Begin VB.Menu mnBarra17 
         Caption         =   "-"
      End
      Begin VB.Menu mnSeleccionarImpresora 
         Caption         =   "Seleccionar &Impresora"
      End
      Begin VB.Menu mnBarra12 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnAlmacen 
      Caption         =   "&Almacen"
      Begin VB.Menu mnDatosGenAlmacen 
         Caption         =   "&Datos Generales"
         Begin VB.Menu mnAlmMarcas 
            Caption         =   "&Marcas"
         End
         Begin VB.Menu mnAlmAlPropios 
            Caption         =   "Almacenes &Propios"
         End
         Begin VB.Menu mnAlmTipoUnidad 
            Caption         =   "Tipos &Unidad"
         End
         Begin VB.Menu mnTiposArticulos 
            Caption         =   "&Tipos Articulos"
         End
         Begin VB.Menu mnAlmFamiliaArticulo 
            Caption         =   "&Familias Artículos"
         End
         Begin VB.Menu mnAlmArticulos 
            Caption         =   "&Artículos"
         End
      End
      Begin VB.Menu mnAlmMovimientosAlm 
         Caption         =   "&Movimientos Almacen"
         Begin VB.Menu mnAlmMovimientos 
            Caption         =   "&Movimientos Almacen"
         End
         Begin VB.Menu mnAlmMovimientosHco 
            Caption         =   "H&istórico Movimientos Almacen"
         End
      End
      Begin VB.Menu mnAlmConsultas 
         Caption         =   "&Consultas"
         Begin VB.Menu mnAlmMovimArticulos 
            Caption         =   "Movimientos A&rticulos"
         End
         Begin VB.Menu mnAlmListMovim 
            Caption         =   "Listado &Movimientos"
         End
         Begin VB.Menu mnAlmListInactivos 
            Caption         =   "Listado Articulos &Inactivos"
         End
         Begin VB.Menu mnAlmListComponentes 
            Caption         =   "Listado Articulos &Componentes"
         End
         Begin VB.Menu mnAlmListValoracion 
            Caption         =   "Listado Valoración &Stocks"
         End
         Begin VB.Menu mnAlmListMaxMin 
            Caption         =   "Inf. Stocks Máximos-Mínimos"
         End
         Begin VB.Menu mnAlmStockFecha 
            Caption         =   "Inf. Stocks a una &Fecha"
         End
      End
      Begin VB.Menu mnAlmInventario 
         Caption         =   "&Inventario"
         Begin VB.Menu mnAlmTomaInven 
            Caption         =   "&Toma de inventario"
         End
         Begin VB.Menu mnAlmEntradaInve 
            Caption         =   "&Entrada existencia real"
         End
         Begin VB.Menu mnAlmListadoInve 
            Caption         =   "&Listado diferencias"
         End
         Begin VB.Menu mnAlmActualizarInve 
            Caption         =   "Actualizar &direrencias"
         End
         Begin VB.Menu mnAlmValoracionInve 
            Caption         =   "&Valoración stocks inventariados"
         End
         Begin VB.Menu mnBarra2 
            Caption         =   "-"
         End
         Begin VB.Menu mnAlmHcoInven 
            Caption         =   "&Histórico inventario"
         End
      End
   End
   Begin VB.Menu mnFacturacion 
      Caption         =   "&Facturación Clientes"
      Begin VB.Menu mnFacDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnFacActividades 
            Caption         =   "Activi&dades"
         End
         Begin VB.Menu mnFacFormasEnvio 
            Caption         =   "Te&xtos Clientes Agrupados"
         End
         Begin VB.Menu mnFacFormasPago 
            Caption         =   "Formas de &Pago"
         End
         Begin VB.Menu mnFacBancosPropios 
            Caption         =   "&Bancos Propios"
         End
         Begin VB.Menu mnFacSituaciones 
            Caption         =   "&Situaciones Especiales"
         End
         Begin VB.Menu mnFacAgentesCom 
            Caption         =   "Agentes &Comerciales"
         End
         Begin VB.Menu mnFacClientesV1 
            Caption         =   "Clientes &Varios"
            Visible         =   0   'False
         End
         Begin VB.Menu mnFacClientes 
            Caption         =   "Cl&ientes"
         End
         Begin VB.Menu mnFacCartas 
            Caption         =   "Tipos de C&artas"
         End
         Begin VB.Menu mnFacIncidencias 
            Caption         =   "&Incidencias"
         End
         Begin VB.Menu mnTarjetas 
            Caption         =   "&Tarjetas"
         End
      End
      Begin VB.Menu mnFacInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnFacInactivos 
            Caption         =   "Clientes Inacti&vos"
         End
         Begin VB.Menu mnFacInfClientes 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu mnFacAltas 
            Caption         =   "&Altas Clientes"
         End
         Begin VB.Menu mnFacEtiqClien 
            Caption         =   "&Etiquetas de clientes"
         End
         Begin VB.Menu mnFacCartaClien 
            Caption         =   "Car&tas a clientes"
         End
      End
      Begin VB.Menu mnTaxitronic 
         Caption         =   "Traspaso TaxiTronic"
      End
      Begin VB.Menu mnHisLlam 
         Caption         =   "Histórico de llamadas"
      End
      Begin VB.Menu mnHisServAso 
         Caption         =   "Mantenimiento Servicios Abonados"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnFacFactClientes 
         Caption         =   "&Facturación a Clientes"
      End
      Begin VB.Menu mnFacFactVarCli 
         Caption         =   "Facturas &Varias a Clientes"
      End
      Begin VB.Menu mnFacCliHcoFact 
         Caption         =   "&Histórico de Facturas"
      End
      Begin VB.Menu mnFacCliReimpr 
         Caption         =   "&Reimprimir Facturas"
      End
      Begin VB.Menu mnFacCliContabilizar 
         Caption         =   "&Contabilizar Facturas"
      End
      Begin VB.Menu mnFacCliRectifica 
         Caption         =   "&Facturas Rectificativas"
      End
      Begin VB.Menu mnBarra 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnFacCliEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnFacCliEstVentaCliente 
            Caption         =   "&Ventas por Cliente"
         End
         Begin VB.Menu mnFacCliDetalleFac 
            Caption         =   "&Detalle facturación"
         End
      End
   End
   Begin VB.Menu mnAdministracion 
      Caption         =   "Facturación Socios"
      Begin VB.Menu mnAdmDatosGen 
         Caption         =   "&Datos Generales"
         Begin VB.Menu mnAdmTrabajadores 
            Caption         =   "&Trabajadores"
         End
         Begin VB.Menu msGesCoche 
            Caption         =   "Vehículos"
         End
         Begin VB.Menu mnGesConduc 
            Caption         =   "&Choferes"
         End
         Begin VB.Menu mnGesSoc 
            Caption         =   "&Socios"
         End
         Begin VB.Menu mnGesUve 
            Caption         =   "&Histórico de Uves"
         End
      End
      Begin VB.Menu mnSocInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnInfVarEtiqSoc 
            Caption         =   "&Etiquetas de Socios"
         End
         Begin VB.Menu mnInfVarCartasSoc 
            Caption         =   "Car&tas a Socios"
         End
      End
      Begin VB.Menu mnFacAlbaran 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnFacEntAlbaran 
            Caption         =   "&Mantenimiento Albaranes"
         End
         Begin VB.Menu mnFacAlbxArtic 
            Caption         =   "Informe &Albaranes por Articulo"
         End
         Begin VB.Menu mnFacHcoAlbaranes 
            Caption         =   "&Histórico Albaranes Anulados"
         End
         Begin VB.Menu mnBarra5 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacPreFacturar 
            Caption         =   "&Previsión Facturación"
         End
         Begin VB.Menu mnFacFacturarAlb 
            Caption         =   "&Facturación de Albaranes"
         End
         Begin VB.Menu mnFacAlbRectifica 
            Caption         =   "Facturas &Rectificativas"
         End
         Begin VB.Menu mnFacHcoFacturas 
            Caption         =   "His&tórico Albaran/Factura"
         End
         Begin VB.Menu mnFacReImpFactu 
            Caption         =   "Re&imprimir Facturas"
         End
         Begin VB.Menu mnTicket 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnFacContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
      End
      Begin VB.Menu mnFacLiquidacion 
         Caption         =   "&Liquidación"
         Begin VB.Menu mnHisServSocios 
            Caption         =   "Mantenimiento Servicios Socios"
         End
         Begin VB.Menu mnFacLiqPdteLiquidar 
            Caption         =   "&Informe Pdte Liquidar"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnFacLiqLiquidacion 
            Caption         =   "&Liquidación"
         End
         Begin VB.Menu mnFacLiqHcoFact 
            Caption         =   "&Histórico Facturas"
         End
         Begin VB.Menu mnFacLiqReimpresion 
            Caption         =   "&Reimprimir Facturas"
         End
         Begin VB.Menu mnFacLiqIntContable 
            Caption         =   "&Contabilizar Facturas"
         End
         Begin VB.Menu mnFacLiqDesFac 
            Caption         =   "&Deshacer Facturación"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnBarra11 
            Caption         =   "-"
         End
         Begin VB.Menu mnFacLiqRetencion 
            Caption         =   "&Retenciones Socio"
         End
      End
      Begin VB.Menu mnFacEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnFacEstVentaCliente 
            Caption         =   "&Ventas por Socio"
         End
         Begin VB.Menu mnFacEstVentaMes 
            Caption         =   "Ventas por &meses"
         End
         Begin VB.Menu mnFacEstVentaFam 
            Caption         =   "Ventas por &familia  /  Artículo"
         End
         Begin VB.Menu mnFacEstDetalleFac 
            Caption         =   "&Detalle facturación"
         End
      End
   End
   Begin VB.Menu mnCompras 
      Caption         =   "&Compras"
      Begin VB.Menu mnComDatosGenerales 
         Caption         =   "Datos &Generales"
         Begin VB.Menu mnComProveedores 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComProveVarios 
            Caption         =   "Proveedores &Varios"
         End
         Begin VB.Menu mnComDirecciones 
            Caption         =   "&Direcciones"
         End
      End
      Begin VB.Menu mnComInfVarios 
         Caption         =   "&Informes Varios"
         Begin VB.Menu mnComInfProve 
            Caption         =   "&Proveedores"
         End
         Begin VB.Menu mnComEtiqProve 
            Caption         =   "&Etiquetas de proveedores"
         End
         Begin VB.Menu mnComCartaProve 
            Caption         =   "&Cartas a Proveedores"
         End
      End
      Begin VB.Menu mnComPreciosDtos 
         Caption         =   "Precios y &Descuentos"
         Begin VB.Menu mnComPreProve 
            Caption         =   "P&recios Proveedor"
         End
         Begin VB.Menu mnComDtosProve 
            Caption         =   "Descuentos Pro&veedor"
         End
      End
      Begin VB.Menu mnComPedidos 
         Caption         =   "&Pedidos"
         Begin VB.Menu mnComPedMant 
            Caption         =   "Mant. &Pedidos Proveedor"
         End
         Begin VB.Menu mnComHcoPedidos 
            Caption         =   "&Histórico Pedidos Anulados"
         End
         Begin VB.Menu mnComPteRecibir 
            Caption         =   "List. &Material pendiente de recibir"
         End
      End
      Begin VB.Menu mnComAlbaranes 
         Caption         =   "&Albaranes"
         Begin VB.Menu mnComAlbMan 
            Caption         =   "&Mant. Albaranes Proveedor"
         End
         Begin VB.Menu mnComHcoAlbaranes 
            Caption         =   "&Histórico Albaranes Anulados"
         End
         Begin VB.Menu mnComPteFacturar 
            Caption         =   "List. &Pendiente de facturar"
         End
         Begin VB.Menu mnBarra7 
            Caption         =   "-"
         End
         Begin VB.Menu mnComFacturar 
            Caption         =   "&Recepción Facturas"
         End
         Begin VB.Menu mnComHcoFacturas 
            Caption         =   "&Histórico Albaran/Factura"
         End
         Begin VB.Menu mnBarra15 
            Caption         =   "-"
         End
         Begin VB.Menu mnComContFactu 
            Caption         =   "&Contabilizar Facturas"
         End
      End
      Begin VB.Menu Barra7 
         Caption         =   "-"
      End
      Begin VB.Menu mnComEstadistica 
         Caption         =   "&Estadística"
         Begin VB.Menu mnComEstComprasxProve 
            Caption         =   "Compras por &Proveedor"
         End
         Begin VB.Menu mnComEstComprasxFam 
            Caption         =   "Compras por &Familia/Artíc."
         End
         Begin VB.Menu mnComEstAlbarxProve 
            Caption         =   "&Albaranes por Proveedor"
         End
      End
   End
   Begin VB.Menu mnpublicidad 
      Caption         =   "Publicidad"
      Begin VB.Menu mnPubliFactuClientes 
         Caption         =   "Facturación &Clientes"
      End
      Begin VB.Menu mnPubliFacRecClientes 
         Caption         =   "Facturas Rectificativas"
      End
      Begin VB.Menu mnhcoFacPubliCli 
         Caption         =   "&Histórico Facturas Clientes"
      End
      Begin VB.Menu mnPubliFactuSocios 
         Caption         =   "Facturación &Socios"
      End
      Begin VB.Menu mnPubliFacRecSocios 
         Caption         =   "Facturas Rectificativas Socios"
      End
      Begin VB.Menu mnhcoFacPubliSoc 
         Caption         =   "Histórico &Facturas Socios"
      End
      Begin VB.Menu mnPubliReimp 
         Caption         =   "&Reimprimir Facturas"
      End
      Begin VB.Menu mnContaFactuPubli 
         Caption         =   "Contabilizar &Facturas"
      End
   End
   Begin VB.Menu mnCuotas 
      Caption         =   "C&uotas"
      Begin VB.Menu mnFactuCuotas 
         Caption         =   "&Generar Facturas Cuotas"
      End
      Begin VB.Menu mnReimpresion 
         Caption         =   "&Reimprimir Facturas"
      End
      Begin VB.Menu mnCuentasHco 
         Caption         =   "&Histórico Facturas"
      End
      Begin VB.Menu mnContaCuotas 
         Caption         =   "&Contabilizar Facturas"
      End
      Begin VB.Menu mnBarraC1 
         Caption         =   "-"
      End
      Begin VB.Menu mnMtoalbaranes 
         Caption         =   "Mantenimiento Albaranes"
      End
      Begin VB.Menu mnPrevFact 
         Caption         =   "&Previsión Facturación"
      End
      Begin VB.Menu mnFacAlb 
         Caption         =   "&Facturación"
      End
      Begin VB.Menu mnBarraC2 
         Caption         =   "-"
      End
      Begin VB.Menu mnFacRecCuo 
         Caption         =   "Facturas Rectificativas"
      End
   End
   Begin VB.Menu mnReparaciones 
      Caption         =   "&Reparaciones"
      Begin VB.Menu mnRepEntReparacion 
         Caption         =   "&Mant.  Reparaciones"
      End
      Begin VB.Menu mnRepControlRep 
         Caption         =   "C&ontrol Reparaciones"
      End
      Begin VB.Menu mnRepNumSerie 
         Caption         =   "Mant. &Nº Serie"
      End
      Begin VB.Menu mnRepMotivosBaja 
         Caption         =   "Motivos &baja equipos"
      End
      Begin VB.Menu mnRepMotivosPend 
         Caption         =   "Motivos &Pend. Rep."
      End
      Begin VB.Menu mnRepHistorico 
         Caption         =   "&Histórico de Reparaciones"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnManServicioAsisTecn 
         Caption         =   "Servicios asistencia técnica"
      End
      Begin VB.Menu mnTiposAveria 
         Caption         =   "Tipos averia"
      End
      Begin VB.Menu mnTrabaRealiz 
         Caption         =   "Trabajos realizados"
      End
      Begin VB.Menu Barra9 
         Caption         =   "-"
      End
      Begin VB.Menu mnRepListRepxDia 
         Caption         =   "Listado Rep. del &Dia"
      End
      Begin VB.Menu mnRepListRepxClien 
         Caption         =   "Listado Rep. por &Cliente"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnRepListFrecuen 
         Caption         =   "F&recuencia de reparaciones"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnEstadisticaReparacionTecnico 
         Caption         =   "Estadística reparaciones técnico"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnListadoReparacionesEfectuadas 
         Caption         =   "Listado reparaciones efectuadas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnRepAlbaranes 
         Caption         =   "Mant. &Albaranes Rep."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnRepPrevFact 
         Caption         =   "Pre&visión Facturación"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnRepFactAlb 
         Caption         =   "&Facturación Reparaciones"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnCRMmenu 
      Caption         =   "CRM"
      Begin VB.Menu mnCRM 
         Caption         =   "Mantenimiento acciones comerciales"
         Index           =   0
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Tipos acciones comerciales"
         Index           =   1
      End
      Begin VB.Menu mnCRM 
         Caption         =   "Generar acciones comerciales"
         Index           =   2
      End
   End
   Begin VB.Menu mnUtilidades 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnVerAvisos 
         Caption         =   "A&visos"
      End
      Begin VB.Menu mnLlamadas 
         Caption         =   "Llamadas"
         Index           =   0
      End
      Begin VB.Menu mnLlamadas 
         Caption         =   "Concepto llamadas"
         Index           =   1
      End
      Begin VB.Menu mnBackUp 
         Caption         =   "&Copia Seguridad local"
      End
      Begin VB.Menu mnRecupFac 
         Caption         =   "&Recuperar facturas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminarFacturas 
         Caption         =   "&Borre Facturas y Movimientos"
      End
      Begin VB.Menu mnRevisarMultibase 
         Caption         =   "Revisar caracteres especiales"
      End
      Begin VB.Menu mnManteneLOG 
         Caption         =   "Acciones realizadas"
      End
      Begin VB.Menu mnEliminarArticulos 
         Caption         =   "Eliminar articulos"
      End
      Begin VB.Menu mnExportarFacturas 
         Caption         =   "Facturación Electrónica"
      End
      Begin VB.Menu mnBarra19 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiBuscar 
         Caption         =   "&Buscar..."
         Begin VB.Menu mnUtiBuscarErrFac 
            Caption         =   "&Errores en Nº Factura clientes"
         End
         Begin VB.Menu mnUtiBuscarPteCon 
            Caption         =   "Facturas pendientes de &contabilizar"
            Begin VB.Menu mnUtiBuscarErrConCli 
               Caption         =   "&Clientes"
            End
            Begin VB.Menu mnUtiBuscarErrConPro 
               Caption         =   "&Proveedores"
            End
         End
      End
      Begin VB.Menu mnBarra20 
         Caption         =   "-"
      End
      Begin VB.Menu mnUtiUsuActivos 
         Caption         =   "&Usuarios activos"
      End
      Begin VB.Menu mnUtiConnActivas 
         Caption         =   "&Conexiones activas"
      End
      Begin VB.Menu mnBarra21 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnSoporte2 
      Caption         =   "&Soporte"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnSoporte 
         Caption         =   "Ayuda"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Enviar Mail"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Web Ariadna Software"
         Index           =   4
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Comprobar version operativa"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnSoporte 
         Caption         =   "Acerca de ..."
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmPpal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrimeraVez As Boolean

Dim TieneEditorDeMenus As Boolean


Private Sub MDIForm_Activate()
Dim b As Boolean

'Dim AvisosPendientes As Boolean
'Formulario Principal
   ' AvisosPendientes = False
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
       ' AvisosPendientes = TieneAvisosPendientes()
    End If
    If Not vParam Is Nothing Then
        If vParam.Modificado Then
          'Poner datos visible del form
           PonerDatosVisiblesForm
           vParam.Modificado = False
        End If
    End If
    '-- Control de si se utilizan servicios o no ( si es que no no se muestra el menú)
    '   el situarlo aqui hace que no haya que salir y entrar en el programa si se
    'vParamAplic.Reparaciones
    
    
    ' *** per als iconos XP ***
'    GetIconsFromLibrary App.Path & "\styles\iconos.dll", 1, 24
'    GetIconsFromLibrary App.Path & "\iconos_BN.dll", 2, 24
'    GetIconsFromLibrary App.Path & "\iconos_OM.dll", 3, 24
'++

'++
    
    
    'Reparaciones
    'mnReparaciones.visible = vParamAplic.Reparaciones
    PuntoDeMenuVisible mnReparaciones, vParamAplic.Reparaciones
    
    
    'De momento:
       
    PuntoDeMenuVisible Me.mnCRMmenu, vParamAplic.TieneCRM
       
       
    
     '-- Descriptores especiales (Vrs 4.0.9)
    If vParamAplic.Descriptores Then
        mnAlmTipoUnidad.Caption = "Formatos"
        mnTiposArticulos.Caption = "Modelos"
        mnAlmFamiliaArticulo.Caption = "Categorias Art."
    End If
    '--
    Screen.MousePointer = vbDefault
End Sub


Public Sub GetIconsFromLibrary(ByVal sLibraryFilePath As String, ByVal Op As Integer, ByVal tam As Integer)
    Dim i As Integer
    Dim tRes As ResType, iCount As Integer
        
    opcio = Op
    tamany = tam
    ghmodule = LoadLibraryEx(sLibraryFilePath, 0, DONT_RESOLVE_DLL_REFERENCES)

    If ghmodule = 0 Then
        MsgBox "Invalid library file.", vbCritical
        Exit Sub
    End If
        
    For tRes = RT_FIRST To RT_LAST
        DoEvents
        EnumResourceNames ghmodule, tRes, AddressOf EnumResNameProc, 0
    Next
    FreeLibrary ghmodule
             
End Sub



Private Sub PuntoDeMenuVisible(ByRef MnPuntoDMenu As Menu, b As Boolean)
    If MnPuntoDMenu.visible Then MnPuntoDMenu.visible = b
    
End Sub


Private Sub MDIForm_Load()
'Formulario Principal

    CargaImagen

    CargaIconosDlls

    PrimeraVez = True
    'Botones
    With Me.Toolbar1
        .ImageList = Me.ImgListPpal
        .Buttons(1).Image = 1   'Articulos
        .Buttons(2).Image = 2   'Movimientos Articulos
        
        .Buttons(5).Image = 3   'Clientes
        .Buttons(6).Image = 4   'Proveedores
        .Buttons(7).Image = 32 'socios
        
        .Buttons(10).Image = 33   'Hco de llamadas
        
        .Buttons(13).Image = 6   'Pedidos Clientes
        .Buttons(14).Image = 7   'Albaranes Clientes
        .Buttons(15).Image = 8   'Hist. Albaranes Clientes (Facturas)

        .Buttons(18).Image = 9   'Pedidos Proveedor
        .Buttons(19).Image = 10   'Albaranes Proveedor
        .Buttons(20).Image = 11   'Facturas Proveedor
        .Buttons(21).Image = 12   'Recepcion Facturas Proveedor
        
        .Buttons(24).Image = 34  ' facturas liquidacion socios
        .Buttons(25).Image = 35  ' facturas publicidad socios
        .Buttons(26).Image = 36  ' facturas cuotas socios
        
        .Buttons(29).Image = 37  ' facturas clientes
        .Buttons(30).Image = 38  ' facturas publicidad clientes
        
        
        .Buttons(32).Image = 16   'Nº Serie
        .Buttons(37).Image = 21 'Cambio de empresa
        
        .Buttons(40).Image = 14 'Salir
    End With
    
    
    
    
    
    LeerEditorMenus
    PonerDatosFormulario False
    
    
End Sub



Private Function CargaIconosDlls()
Dim TamanyoImgComun As Integer
    
    
    imgListComun1.ListImages.Clear
    imgListComun_BN.ListImages.Clear
    imgListComun_OM.ListImages.Clear
    
    TamanyoImgComun = 24
    
    imgListComun1.ImageHeight = TamanyoImgComun
    imgListComun1.ImageWidth = TamanyoImgComun
    GetIconsFromLibrary App.Path & "\styles\iconos.dll", 5, TamanyoImgComun  'antes icolistcon

    
    imgListComun_BN.ImageHeight = TamanyoImgComun
    imgListComun_BN.ImageWidth = TamanyoImgComun
    GetIconsFromLibrary App.Path & "\styles\iconos_BN.dll", 2, TamanyoImgComun
  
    imgListComun_OM.ImageHeight = TamanyoImgComun
    imgListComun_OM.ImageWidth = TamanyoImgComun
    GetIconsFromLibrary App.Path & "\styles\iconos_OM.dll", 3, TamanyoImgComun
    
    imgListComun16.ImageHeight = 16
    imgListComun16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\iconos.dll", 1, 16
    
    GetIconsFromLibrary App.Path & "\styles\iconos_BN.dll", 2, 16
    GetIconsFromLibrary App.Path & "\styles\iconos_OM.dll", 3, 16

End Function

Private Sub CargaImagen()
    On Error Resume Next
    If vParamAplic.Cooperativa = 0 Then
        Me.Picture = LoadPicture(App.Path & "\arifon2.dat")
    Else
        Me.Picture = LoadPicture(App.Path & "\arifon3.dat")
    End If
    If Err.Number <> 0 Then
        Me.Picture = LoadPicture()
        Err.Clear
    End If
End Sub




Private Sub PonerDatosFormulario(DesdeCambiarEmpresa As Boolean)
Dim Config As Boolean


    If Not DesdeCambiarEmpresa Then
        Config = (vEmpresa Is Nothing) Or (vParam Is Nothing) Or (vParamAplic Is Nothing)
    
        If Config Then HabilitarSoloPrametros_o_Empresas False
    End If
    
    'FijarConerrores
    CadenaDesdeOtroForm = ""

    'Poner datos visible del form
    PonerDatosVisiblesForm
    
    'Habilitar/Deshabilitar entradas del menu segun el nivel de usuario
    PonerMenusNivelUsuario

    'Si no hay carpeta interaciones, no habra integraciones
'    Me.mnComprobarPendientes.Enabled = vConfig.Integraciones <> ""


    'Habilitar
    If DesdeCambiarEmpresa Then
        ReestablecerMenus
        HabilitarSoloPrametros_o_Empresas True
    End If
    

    'Si tiene editor de menus
    If TieneEditorDeMenus Then PoneMenusDelEditor
    
    'Comprobar que los iconos de la barra su correspondiente
    'entrada de menu esta habilitada sino desabilitar
    PoneBarraMenus
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'Formulario Principal
Dim Cad As String

    'Alguna cosilla antes de cerrar. Eliminar bloqueos
    Cad = "Delete from zbloqueos where codusu = " & vUsu.Codigo
    conn.Execute Cad

    'Elimnar bloquo BD
    Set vUsu = Nothing
    Set vConfig = Nothing
    Set vEmpresa = Nothing
    
    Set vParam = Nothing
    Set vParamAplic = Nothing
    
    
    TerminaBloquear
    
    'cerrar las conexiones
    conn.Close
    CerrarConexionConta

End Sub

Private Sub mnAdmTrabajadores_Click()
    frmAdmTrabajadores.Show vbModal
End Sub

Private Sub mnAlbaranesB_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALZ"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnAlmActualizarInve_Click()
    AbrirListado (14)
End Sub

Private Sub mnAlmAlPropios_Click()
    frmAlmAlPropios.Show vbModal
End Sub

Private Sub mnAlmArticulos_Click()
    frmAlmArticulos.Show vbModal
End Sub

Private Sub mnAlmEntradaInve_Click()
    frmAlmInventario.Show vbModal
End Sub

Private Sub mnAlmFamiliaArticulo_Click()
    frmAlmFamiliaArticulo.Show vbModal
End Sub


Private Sub mnAlmHcoInven_Click()
    frmAlmHcoInven.Show vbModal
End Sub

Private Sub mnAlmListadoInve_Click()
    AbrirListado (13)
End Sub

Private Sub mnAlmListComponentes_Click()
'Informe de articulos q estan compuestos de otros articulos
    AbrirListado (11)
End Sub

Private Sub mnAlmListInactivos_Click()
    AbrirListado (15)
End Sub

Private Sub mnAlmListMaxMin_Click()
'Informe de Stocks Maximos y Minimos
    AbrirListado (18)
End Sub

Private Sub mnAlmListMovim_Click()
    AbrirListado (9)
End Sub

Private Sub mnAlmListValoracion_Click()
    AbrirListado (17)
End Sub

Private Sub mnAlmMarcas_Click()
    frmAlmMarcas.Show vbModal
End Sub

Private Sub mnAlmMovimArticulos_Click()
    frmAlmMovimArticulos.Show vbModal
End Sub

Private Sub mnAlmMovimientos_Click()
    frmAlmMovimientos.EsHistorico = False
    frmAlmMovimientos.hcoCodMovim = -1 'No carga el form al abrir
    frmAlmMovimientos.Show vbModal
End Sub

Private Sub mnAlmMovimientosHco_Click()
    frmAlmMovimientos.EsHistorico = True
    frmAlmMovimientos.hcoCodMovim = -1
    frmAlmMovimientos.Show vbModal
End Sub

Private Sub mnAlmStockFecha_Click()
'Informe de Stocks a una Fecha
    AbrirListado (19)
End Sub

Private Sub mnAlmTipoUnidad_Click()
    frmAlmTipoUnidad.Show vbModal
End Sub

Private Sub mnAlmTomaInven_Click()
    AbrirListado (12)
End Sub

Private Sub mnAlmValoracionInve_Click()
    AbrirListado (16)
End Sub

Private Sub mnBackUp_Click()
'Copia de seguridad de toda la base de datos
    frmBackUP.Show vbModal
End Sub

Private Sub mnBorrarAvisosCerrados_Click()
    AbrirListado 83
End Sub



Private Sub mnCambioEmpresa_Click()
    Dim AntUSU As Usuario

    If Not (Me.ActiveForm Is Nothing) Then
        MsgBox "Cierre todas las ventanas para poder cambiar de usuario", vbExclamation
        Exit Sub
    End If

    'Borramos temporal
    conn.Execute "Delete from zbloqueos where codusu = " & vUsu.Codigo


    CadenaDesdeOtroForm = vUsu.Login & "|" & vUsu.PasswdPROPIO & "|"
    
'    Set AntUSU = vUsu
'    Set vUsu = Nothing
    frmLogin.Show vbModal
'    If vUsu Is Nothing Then
'        Set vUsu = AntUSU
'        Set AntUSU = Nothing
'        Exit Sub
'    End If

    Screen.MousePointer = vbHourglass
    'Cerramos la conexion
    conn.Close
    ConnConta.Close


    'Abre la conexión a BDatos:AriTaxi
    If AbrirConexion() = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    Else
        Set vParam = Nothing
        Set vParamAplic = Nothing
        'Carga Parametros Generales y Contables de la empresa
        LeerParametros
    End If


    'Abrir conexión a la BDatos de Contabilidad para acceder a
    'Tablas: Cuentas, Tipos IVA
    If AbrirConexionConta(False) = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos de Contabilidad. ", vbCritical
        End
    End If

    
    
    Set vEmpresa = Nothing
    'LeerEmpresaParametros
    
     'Carga los Datos Básicos de la empresa
    LeerDatosEmpresa
    
    
    'Carga los Niveles de cuentas de Contabilidad de la empresa
    LeerNivelesEmpresa
    
'    PonerDatosFormulario
    PonerDatosVisiblesForm

    'Ponemos primera vez a false
    PonerDatosFormulario True
    PrimeraVez = True
    MDIForm_Activate

    CargaImagen

    Screen.MousePointer = vbDefault
End Sub

Private Sub mnComAlbMan_Click()
'Mantenimiento de Albaranes a Proveedor
    frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmComEntAlbaranes.EsHistorico = False
    frmComEntAlbaranes.Show vbModal
End Sub

Private Sub mnComCartaProve_Click()
'Cartas a proveedores
     AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
End Sub

Private Sub mnComContFactu_Click()
'Contabilizar Facturas
    AbrirListado (224) 'Para pedir datos
End Sub

Private Sub mnComDirecciones_Click()
    frmComDirecciones.Show vbModal
End Sub

Private Sub mnComDtosProve_Click()
    frmComDtosFamMarca.Show vbModal
End Sub

Private Sub mnComEstAlbarxProve_Click()
'Listado de compras por proveedor
    AbrirListadoOfer (312)
End Sub

Private Sub mnComEstComprasxFam_Click()
'Listado de compras por Familia
    AbrirListadoOfer (311)
End Sub

Private Sub mnComEstComprasxProve_Click()
'Listado de compras por proveedor
    AbrirListadoOfer (310)
End Sub

Private Sub mnComEtiqProve_Click()
'Etiquetas de proveedores
     AbrirListadoOfer (305) '305: Informe Etiquetas de Proveedores
End Sub

Private Sub mnComFacturar_Click()
    frmComFacturar.Show vbModal
End Sub

Private Sub mnComHcoAlbaranes_Click()
'Historico albaranes de compras a proveedores
    frmComEntAlbaranes.EsHistorico = True
    frmComEntAlbaranes.Show vbModal
End Sub

Private Sub mnComHcoFacturas_Click()
    frmComHcoFacturas.hcoCodMovim = ""
    frmComHcoFacturas.Show vbModal
End Sub

Private Sub mnComHcoPedidos_Click()
    frmComEntPedidos.MostrarDatos = ""
    frmComEntPedidos.EsHistorico = True
    frmComEntPedidos.Show vbModal
End Sub

Private Sub mnComInfProve_Click()
'Informe de Proveedores
    AbrirListado (58)   ': Informe Proveedores
End Sub

Private Sub mnComPedMant_Click()
'Mnatenimiento de Pedidos de compras
    frmComEntPedidos.MostrarDatos = ""
    frmComEntPedidos.EsHistorico = False
    frmComEntPedidos.Show vbModal
End Sub

Private Sub mnComPreProve_Click()
    frmComPreciosProv.Show vbModal
End Sub

Private Sub mnComProveedores_Click()
'Compras. Mantenimiento de Proveedores
    frmComProveedores.Show vbModal
End Sub


Private Sub mnComProveVarios_Click()
'Proveedores varios
    frmComProveV.Show vbModal
End Sub

Private Sub mnComPteFacturar_Click()
'Listado de Albaranes pendientes de Factura
    AbrirListadoOfer (308) '308: List. Albaranes pte facturar
End Sub

Private Sub mnComPteRecibir_Click()
'Listado de material pendiente de recibir
    AbrirListadoOfer (307) '307: List. Materia pte recibir
End Sub

Private Sub mnConfManteUsuarios_Click()
'Mantenimiento de Usuarios

      frmMantenusu.Show vbModal
      
End Sub

Private Sub mnConfParamAplic_Click()
'Parametros de la Aplicación
    Screen.MousePointer = vbHourglass
    Load frmConfParamAplic
    frmConfParamAplic.Show vbModal
    
End Sub



Private Sub mnConfParamGenerales_Click()
    
'Parametros generales de la Empresa
    frmConfParamGral.Show vbModal
End Sub



Private Sub mnConfParamRpt_Click()
'Parametros para informes de Crystal Report
    frmConfParamRpt.Show vbModal
End Sub

Private Sub mnContaCuotas_Click()
    frmCuotasContaFac.Show vbModal
End Sub

Private Sub mnContaFactuPubli_Click()
    frmPubliContaFac.Show vbModal
End Sub

Private Sub mnConTMovimiento_Click()
'Mantenimientos de los tipos de movimientos
    frmConfTipoMov.Show vbModal
End Sub


Private Sub mnCRM_Click(Index As Integer)
    
        Select Case Index
        Case 0
            frmCRMMto.DesdeElCliente = 0 'No clien
            frmCRMMto.TipoPredefinido = 0   'Ninguno
            frmCRMMto.Show vbModal
            
        Case 1
            frmCRMtipos.Show vbModal
        
        Case 2
            frmCRMVarios.Opcion = 0
            frmCRMVarios.Show vbModal
            
        
        End Select
        
End Sub


Private Sub mnCuentasHco_Click()
    frmCuotasHcoFacturas.Show vbModal
End Sub

Private Sub mnEliminarArticulos_Click()
    frmVarios.Opcion = 1
    frmVarios.Show vbModal
End Sub

Private Sub mnEliminarFacturas_Click()
    AbrirListado 97
End Sub


Private Sub mnEstadisticaReparacionTecnico_Click()
    AbrirListado2 2
End Sub

Private Sub mnEtiqEstanteria_Click()
    AbrirListado 94
End Sub

Private Sub mnEtiquetasBultos_Click()
'Listado de etiquetas de los bultos
    AbrirListado 95
End Sub

Private Sub mnExportarFacturas_Click()
    frmExportarFacturas.Show vbModal
End Sub

Private Sub mnFacActividades_Click()
    frmFacActividades.Show vbModal
End Sub

Private Sub mnFacAgentesCom_Click()
    frmFacAgentesCom.Show vbModal
End Sub

Private Sub mnFacAlb_Click()
'Facturacion de Albaranes de Ventas
    frmListadoPed.CodClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnFacAlbRectifica_Click()
'Facturas Rectificativas
    'Abre el formulario de Albaranes para introducir el Albaran Rectificativo
    'y desde este generar la Factura Rectificativa
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ART"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacAlbxArtic_Click()
'Informe de Albaranes por Articulo
    AbrirListadoPed (49)
End Sub

Private Sub mnFacAltas_Click()
'Informe de Altas de Nuevos Clientes
    AbrirListadoOfer (48) '48: Informes Altas Clientes
End Sub

Private Sub mnFacBancosPropios_Click()
    frmFacBancosPropios.Show vbModal
End Sub

Private Sub mnFacCartaClien_Click()
'Cartas a clientes
     AbrirListadoOfer (91) '91: Informe Cartas a Clientes
End Sub

Private Sub mnFacCartas_Click()
'Mantenimiento de Cartas
    frmFacCartasOferta.Show vbModal
End Sub


Private Sub mnFacCliContabilizar_Click()
    frmFCliContaFac.Show vbModal
End Sub

Private Sub mnFacCliDetalleFac_Click()
'Detalle facturacion clientes
     AbrirListadoOfer (232)
End Sub

Private Sub mnFacClientes_Click()
'Mantenimiento de Clientes
    frmFacClientes.Show vbModal
End Sub

Private Sub mnFacClientesV1_Click()
'Mantenimiento de Clientes Varios
    frmFacClientesV.Show vbModal
End Sub



Private Sub mnFacCliEstVentaCliente_Click()
'Estadistica Ventas por cliente
    AbrirListadoPed (230)
    BorrarTempInformes
End Sub

Private Sub mnFacCliHcoFact_Click()
    frmFCliHcoFac.Show vbModal
End Sub

Private Sub mnFacCliRectifica_Click()
    frmFCliRectif.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFCliRectif.hcoCodTipoM = "ARN" ' albaran rectificativo de cliente
    frmFCliRectif.EsHistorico = False
    frmFCliRectif.RecuperarFactu = False
    frmFCliRectif.Show vbModal
End Sub

Private Sub mnFacCliReimpr_Click()
    frmFCliReImp.Show vbModal
End Sub

Private Sub mnFacContFactu_Click()
'Contabilizar Facturas
    AbrirListado (223) 'Para pedir datos
End Sub

Private Sub mnFacEntAlbaran_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALV"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub


Private Sub mnFacEstDetalleFac_Click()
'Detalle facturacion clientes
     AbrirListadoOfer (231)
End Sub

Private Sub mnFacEstMargenVtas_Click()
    'Estadistica margen ventas por artículo
        AbrirListado (246)
End Sub

Private Sub mnFacEstVentaAgente_Click()
    'Ventas por agente
    AbrirListado2 16
End Sub

Private Sub mnFacEstVentaCliente_Click()
'Estadistica Ventas por cliente
    AbrirListadoPed (227)
    BorrarTempInformes
End Sub

Private Sub mnFacEstVentaFam_Click()
'Listado de estadistica ventas por familia de articulo
    AbrirListadoOfer (230)
End Sub

Private Sub mnFacEstVentaMes_Click()
'Estadistica Ventas por Meses
    AbrirListadoPed (229)
    
End Sub

Private Sub mnFacEstVentaTraba_Click()
'Estadistica Ventas por Trabajador
    AbrirListadoPed (228)
End Sub

Private Sub mnFacEtiqClien_Click()
'Etiquetas de clientes
     AbrirListadoOfer (90) '90: Informe Etiquetas de Clientes
End Sub

Private Sub mnFacFactClientes_Click()
    frmFCliFacturac.Show vbModal
End Sub

Private Sub mnFacFacturarAlb_Click()
'Facturacion de Albaranes de Ventas
    frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnFacFactVarCli_Click()
    frmFCliFactuVar.Show vbModal
End Sub

Private Sub mnFacFormasPago_Click()
    frmFacFormasPago.Show vbModal
End Sub



Private Sub mnFacHcoAlbaranes_Click()
'Histórico de Albaranes eliminados
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALV"
    frmFacEntAlbaranes.EsHistorico = True
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacHcoFacturas_Click()
'Histórico de Facturas
    frmFacHcoFacturas2.hcoCodMovim = ""
    frmFacHcoFacturas2.publicidad = False
    frmFacHcoFacturas2.Show vbModal
End Sub


Private Sub mnFacInactivos_Click()
'Informe de Clientes Inactivos
    AbrirListadoOfer (46) '46: Informes Clientes Inactivos
End Sub

Private Sub mnFacIncidencias_Click()
    frmIncidencias.Show vbModal
End Sub

Private Sub mnFacInfClientes_Click()
'Informe de Clientes
    AbrirListadoOfer (47) '47: Informes Clientes
End Sub

Private Sub mnFacOfertas_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
    
    Select Case Index
    Case 0, 5
            'Private Sub mnFacEntOfertas_Click()

    Case 1
            'Private Sub mnFacGrupoPlant_Click()
            'Mantenimiento de Grupos de Plantillas
            'frmFacGrupoPlantilla.Show vbModal
    
    Case 2
            'Private Sub mnFacPlantillas_Click()
            'Mantenimiento de Plantillas
        
    Case 3
            ' Private Sub mnFacOfeEfectuadas_Click()
            'Listado de Ofertas Efectuadas
        AbrirListadoOfer (34) '34: Informe Ofertas Efectuadas
    
        
        
    'case 4  'Es la barra separadora
    
    Case 6
        
            'Private Sub mnFacTrasHist_Click()
            'Traspaso de Ofertas a las tablas de Historico
        frmListadoOfer.OpcionListado = 36
        AbrirListadoOfer (36) 'NO IMPRIME LISTADO, hace traspaso de Ofertas de la tabla (scapre) a (schpre)

    
    End Select
End Sub

Private Sub mnFacPedidos_Click(Index As Integer)
    'Estan todos agrupados bajo el mismo mn
  
    Select Case Index
    
    'Case 2  es la barra de separacion
    
    Case 3
        'Confirmar pedido   mnFacConfirmPed_Click
        AbrirListadoOfer (40)
        
    Case 8
'       frmFacConsultaPrecios.Show vbModal
    Case 9
        frmVarios.Opcion = 2
        frmVarios.Show vbModal
    End Select
End Sub

Private Sub mnFacLiqDesFac_Click()
    frmLiqDeshacerFac.Show vbModal
End Sub

Private Sub mnFacLiqHcoFact_Click()
    frmLiqHcoFacSoc.Show vbModal
End Sub

Private Sub mnFacLiqIntContable_Click()
    frmLiqContaFac.Show vbModal
End Sub

Private Sub mnFacLiqLiquidacion_Click()
    frmLiqLiquidaSoc.Show vbModal
End Sub

Private Sub mnFacLiqPdteLiquidar_Click()
    frmLiqPdteLiquida.Show vbModal
End Sub

Private Sub mnFacLiqReimpresion_Click()
    frmLiqReImp.Show vbModal
End Sub

Private Sub mnFacLiqRetencion_Click()
    frmLiqRetencion.Show vbModal
End Sub

Private Sub mnFacPreFacturar_Click()
' Previsión Facturacion de Albaranes
    frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO
End Sub

Private Sub mnFacRecCuo_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ARC"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnFacReImpFactu_Click()
'Reimprimir Factuas ya contabilizadas
    AbrirListadoOfer 226
End Sub

Private Sub mnFacSituaciones_Click()
    frmFacSituaciones.Show vbModal
End Sub

Private Sub mnFactuCuotas_Click()
    frmCuotasFac.Show vbModal
End Sub

Private Sub mnFacturarPresupuestos_Click()
        frmListadoPed.CodClien = "ALZ" 'utilizamos esta vble para pasarle el tipo de movimiento
        AbrirListadoPed (52)
End Sub

Private Sub mnFacFormasEnvio_Click()
    frmFacFormasEnvio.Show vbModal
End Sub

Private Sub mnGesConduc_Click()
    frmGesConduc.Show vbModal
End Sub

Private Sub mnGesSoc_Click()
    frmGesSocios.Show vbModal
End Sub

Private Sub mnGesUve_Click()
    frmGesHcoUves.Show vbModal
End Sub

Private Sub mnhcoFacPubliCli_Click()
'Histórico de Facturas
    frmPubliHcoFacCli.Show vbModal
End Sub

Private Sub mnhcoFacPubliSoc_Click()
    frmPubliHcoFacSoc.Show vbModal
End Sub

Private Sub mnHisLlam_Click()
    Select Case vParamAplic.Cooperativa
        Case 0
            frmGesHisLlam.Show vbModal
        Case 1
            frmGesHisLlamVIP.Show vbModal
    End Select
End Sub

Private Sub mnHisServAso_Click()
    frmGesServAbonados.Show vbModal
End Sub

Private Sub mnHisServSocios_Click()
    frmGesServSocios.Show vbModal
End Sub

Private Sub mnInfVarCartasSoc_Click()
    AbrirListadoOfer 191
End Sub

Private Sub mnInfVarEtiqSoc_Click()
    AbrirListadoOfer 190
End Sub

Private Sub mnListadoReparacionesEfectuadas_Click()
    AbrirListado2 1
End Sub

Private Sub mnLlamadas_Click(Index As Integer)
    Select Case Index
    Case 0
        frmLlamadas.Show vbModal
        
    Case 1
        frmLlamadasTipo.Show vbModal
    End Select
End Sub

Private Sub mnManServicioAsisTecn_Click()
    frmManSat.Show vbModal
End Sub



Private Sub mnManteneLOG_Click()
    Screen.MousePointer = vbHourglass
    Load frmLog
    DoEvents
    frmLog.Show vbModal
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnMtoalbaranes_Click()
    frmFacEntAlbaranes.hcoCodTipoM = "ALS"
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnPrevFact_Click()
' Previsión Facturacion de Albaranes
    frmListadoPed.CodClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO
End Sub

Private Sub mnPubliFacRecClientes_Click()
'    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
'    frmFacEntAlbaranes.hcoCodTipoM = "ARP"
'    frmFacEntAlbaranes.EsHistorico = False
'    frmFacEntAlbaranes.RecuperarFactu = False
'    frmFacEntAlbaranes.Show vbModal
    frmFCliRectif.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFCliRectif.hcoCodTipoM = "ARP" ' albaran rectificativo de cliente
    frmFCliRectif.EsHistorico = False
    frmFCliRectif.RecuperarFactu = False
    frmFCliRectif.Show vbModal
End Sub

Private Sub mnPubliFacRecSocios_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ARQ"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnPubliFactuClientes_Click()
    frmPubliFacCli.Show vbModal
End Sub

Private Sub mnPubliFactuSocios_Click()
    frmPubliFacSoc.Show vbModal
End Sub

Private Sub mnPubliReimp_Click()
    frmPubliReImp.Show vbModal
End Sub

Private Sub mnRecupFac_Click()
'recuperar facturas
    'abrimos albaranes de mostrador
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALM"
    'le indicamos q estamos recuperando facturas
    frmFacEntAlbaranes.RecuperarFactu = True
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnReimpresion_Click()
    frmCuotasReImp.Show vbModal
End Sub

Private Sub mnRepAlbaranes_Click()
    frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
    frmFacEntAlbaranes.hcoCodTipoM = "ALR"
    frmFacEntAlbaranes.EsHistorico = False
    frmFacEntAlbaranes.RecuperarFactu = False
    frmFacEntAlbaranes.Show vbModal
End Sub

Private Sub mnRepControlRep_Click()
'Control de Reparaciones (para los Tecnicos)
    frmRepEntReparaciones.EntradaEquipo = ""
    frmRepEntReparaciones.ControlRep = True
    frmRepEntReparaciones.EsHistorico = False
    frmRepEntReparaciones.Show vbModal
End Sub

Private Sub mnRepEntReparacion_Click()
'Mantenimiento de Reparaciones
    frmRepEntReparaciones.EntradaEquipo = ""
    frmRepEntReparaciones.ControlRep = False
    frmRepEntReparaciones.EsHistorico = False
    frmRepEntReparaciones.Show vbModal
End Sub

Private Sub mnRepFactAlb_Click()
'Facturacion de Albaranes de Reparacion
    frmListadoPed.CodClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (52)
End Sub

Private Sub mnRepHistorico_Click()
'Historico de las reparaciones
    frmRepEntReparaciones.EntradaEquipo = ""
    frmRepEntReparaciones.ControlRep = False
    frmRepEntReparaciones.EsHistorico = True
    frmRepEntReparaciones.Show vbModal
End Sub


Private Sub mnRepListAvisosPtes_Click()
'Listado de avisos de averias de clientes pendientes
    AbrirListado (409)
End Sub

Private Sub mnRepListFrecuen_Click()
'Listado de Frecuencia de Reparaciones
    AbrirListado (406)
End Sub

Private Sub mnRepListRepxClien_Click()
'Listado de las Reparaciones por cliente
    AbrirListado (64)
End Sub

Private Sub mnRepListRepxDia_Click()
'Listado de las Reparaciones del dia
    AbrirListado (63)
End Sub

Private Sub mnRepMotivosBaja_Click()
'Motivos baja equipos
    frmRepMotivosBaja.Show vbModal
End Sub

Private Sub mnRepMotivosPend_Click()
'Motivos Pendientes Reparar
    frmRepMotivosPend.Show vbModal
End Sub

Private Sub mnRepNumSerie_Click()
'Mantenimiento de Nºs de Serie
    frmRepNumSerie2.Show vbModal
End Sub

Private Sub mnRepPrevFact_Click()
' Previsión Facturacion de Albaranes de Reparacion
    frmListadoPed.CodClien = "ALR" 'utilizamos esta vble para pasarle el tipo de movimiento
    AbrirListadoPed (50) 'NO IMPRIME LISTADO

End Sub

Private Sub mnRevisarMultibase_Click()
    AbrirListado2 3
End Sub

'Private Sub mnPedirPwd_Click()
'Dim Anterior As Boolean
'
'    Anterior = Me.mnPedirPwd.Checked
'    vConfig.PedirPasswd = Not Anterior
'    If vConfig.Grabar = 1 Then
'        Me.mnPedirPwd.Checked = Anterior
'    Else
'        Me.mnPedirPwd.Checked = Not Anterior
'    End If
'End Sub


Private Sub mnSeleccionarImpresora_Click()
    Screen.MousePointer = vbHourglass
    Me.CommonDialog1.ShowPrinter
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnSociosProveedores_Click(Index As Integer)
    Select Case Index
    Case 0
        'Cambiar precios proveedores /socios
         AbrirListado2 7
         
    Case 1
        'Liquidacion SOCIOS
        AbrirListado2 8
        
    Case 2
        'Impresion facturas proveedores
        AbrirListado2 9
        
    End Select
End Sub

Private Sub mnSoporte_Click(Index As Integer)
    Select Case Index
    Case 4
       
        Screen.MousePointer = vbHourglass
        LanzaHome ("websoporte")
        Screen.MousePointer = vbDefault

    Case 7
        'Acerca de
        Screen.MousePointer = vbDefault
        frmMensajes.OpcionMensaje = 3
        frmMensajes.Show vbModal
    End Select
    
End Sub

Private Sub mnTarjetas_Click()
    frmTarjetas.Show vbModal
End Sub

Private Sub mnTaxitronic_Click()
    frmGesTraspaso.Show vbModal
End Sub

Private Sub mnTicket_Click(Index As Integer)
    
    If Index > 0 Then AbrirListado2 12 + Index

    
End Sub

Private Sub mnTiposArticulos_Click()
    frmAlmTipoArticulo.Show vbModal
End Sub

Private Sub mnSalir_Click()
    End
End Sub





Private Sub mnTiposAveria_Click()
    frmtipave.Show vbModal
End Sub

Private Sub mnTPVcierreCaja_Click()
'Abre el informe de cierre de caja del dia en el TPV
    AbrirListadoOfer (240)
End Sub

Private Sub mnTrabaRealiz_Click()
    frmManTraReali.Show vbModal
End Sub


Private Sub mnUtiBuscarErrConCli_Click()
'Facturas pendientes de contabilizar (CLIENTES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 6
    frmUtilidades.Show vbModal
End Sub

Private Sub mnUtiBuscarErrConPro_Click()
'Facturas pendientes de contabilizar (PROVEEDORES)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 7
    frmUtilidades.Show vbModal
End Sub


Private Sub mnUtiBuscarErrFac_Click()
'Buscar errores en nº de factura (solo en facturas de clientes)
    Screen.MousePointer = vbHourglass
    frmUtilidades.Opcion = 5
    frmUtilidades.Show vbModal
End Sub



Private Sub mnUtiConnActivas_Click()
'ver las conexiones a donde apuntan
Dim Cad As String
 
    
    MostrarCadenasConexion
End Sub

Private Sub mnUtiUsuActivos_Click()
'Muestra si hay otros usuarios conectados a la Gestion
Dim Sql As String
Dim i As Integer

    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        i = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            Sql = RecuperaValor(CadenaDesdeOtroForm, i)
            If Sql <> "" Then Me.Tag = Me.Tag & "    - " & Sql & vbCrLf
            i = i + 1
        Loop Until Sql = ""
        MsgBox Me.Tag, vbExclamation
    Else
        MsgBox "Ningun usuario, además de usted, conectado a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf, vbInformation
    End If
    CadenaDesdeOtroForm = ""
End Sub



Private Sub mnVentasPorProveedor_Click()
        AbrirListado2 6
End Sub

Private Sub mnVerAvisos_Click()
    If TieneAvisosPendientes Then
        frmAlertas.Show vbModal
    Else
        MsgBox "No hay avisos para mostrar", vbInformation
    End If
End Sub







Private Sub msGesCoche_Click()
    frmGesVehic.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
    Case 1 'Mantenimiento de Artículos
        mnAlmArticulos_Click
    Case 2 'Movimientos Articulos
        mnAlmMovimArticulos_Click
        
    Case 5 'Mantenimiento Clientes
        mnFacClientes_Click
    Case 6 'Mantenimiento Proveedores
        mnComProveedores_Click
    Case 7 'mantenimiento de socios
        mnGesSoc_Click
        
    Case 10 ' Hco de llamadas
         mnHisLlam_Click
        
    Case 13 'Pedidos a Clientes
        'mnFacEntPedidos_Click
        mnFacPedidos_Click 0
    Case 14 'Albaranes a Clientes
        mnFacEntAlbaran_Click
    Case 15 'Hist. Albaranes (Facturas)
        mnFacHcoFacturas_Click
        
    Case 18 'Pedidos de Proveedores
        mnComPedMant_Click
    Case 19 'Albaranes de Proveedores
        mnComAlbMan_Click
    Case 20 'Facturas de Proveedores
        mnComHcoFacturas_Click
    Case 21 'Recepcion Fact. Prove
        If Me.mnComFacturar.visible And Me.mnComFacturar.Enabled Then mnComFacturar_Click
        
    Case 24 'Hco.Facturas Liquidacion socio
        frmLiqHcoFacSoc.Show vbModal
    Case 25 'Hco.Facturas Publicidad Socio
        frmPubliHcoFacSoc.Show vbModal
    Case 26 'Hco.Facturas Cuota Socio
        frmCuotasHcoFacturas.Show vbModal
    
    Case 29 'Hco.Facturas Cliente
        frmFCliHcoFac.Show vbModal
    Case 30 'Hco.Facturas Publicidad Cliente
        frmPubliHcoFacCli.Show vbModal
    
    Case 32 'Nº Serie
        mnRepNumSerie_Click
        
    Case 28
        'Consulta precio articulo
        mnFacPedidos_Click 8
        
    Case 37
        'cambiar empresa
        mnCambioEmpresa_Click
        
    Case 40 'Salir
        mnSalir_Click
    End Select
End Sub


Private Sub PonerDatosVisiblesForm()
'Escribe texto de la barra de la aplicación
Dim Cad As String
    Cad = UCase(Mid(Format(Now, "dddd"), 1, 1)) & Mid(Format(Now, "dddd"), 2)
    Cad = Cad & ", " & Format(Now, "d")
    Cad = Cad & " de " & Format(Now, "mmmm")
    Cad = Cad & " de " & Format(Now, "yyyy")
    Cad = "    " & Cad & "    "
    Me.StatusBar1.Panels(5).Text = Cad
    If vEmpresa Is Nothing Then
        Caption = "AriTaxi" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & "   Usuario: " & vUsu.Nombre & " FALTA CONFIGURAR"
        'Panel con el nombre de la empresa
        Me.StatusBar1.Panels(2).Text = "Falta configurar"
    Else
        Caption = "AriTaxi" & " ver. " & App.Major & "." & App.Minor & "." & App.Revision & "   -  " & vEmpresa.nomempre & "  -    Usuario: " & vUsu.Nombre
        Me.StatusBar1.Panels(2).Text = "Empresa:   " & vEmpresa.nomempre & "               Código: " & vEmpresa.codempre
    End If
End Sub


Private Sub HabilitarSoloPrametros_o_Empresas(Habilitar As Boolean)
Dim T As Control
Dim Cad As String

    
    For Each T In Me
        Cad = T.Name
        If Mid(T.Name, 1, 2) = "mn" Then
            If LCase(Mid(T.Caption, 1, 1)) <> "-" Then T.Enabled = Habilitar
        End If
    Next
    Me.Toolbar1.Enabled = Habilitar
    Me.Toolbar1.visible = Habilitar
    Me.mnConfParamAplic = True
    Me.mnConfParamGenerales = True

    Me.mnSalir.Enabled = True
    Me.mnCambioEmpresa.Enabled = True
End Sub

'-------------------------------------
'Pondremos todos los que menus a visibles. Y luego ya , en f
Private Sub ReestablecerMenus()
Dim T
      For Each T In Me
        If Mid(T.Name, 1, 2) = "mn" Then T.visible = True
    Next
End Sub

Private Sub PonerMenusNivelUsuario()
Dim b As Boolean

    b = (vUsu.Nivel = 0) Or (vUsu.Nivel = 1)  'Administradores y root

    Me.mnConfParamAplic = b
    mnConfManteUsuarios = b
    
    mnUsuarios.Enabled = b
    mnNuevaEmpresa.Enabled = b
    mnPedirPwd.Enabled = b
    Me.mnUtiConnActivas.Enabled = (vUsu.Nivel = 0) 'solo para root
    

    b = vUsu.Nivel = 3  'Es usuario de consultas
    If b Then
        'Inventario
        Me.mnAlmTomaInven.Enabled = False
        Me.mnAlmEntradaInve.Enabled = False
        Me.mnAlmActualizarInve.Enabled = False
        Me.mnAlmListadoInve.Enabled = False
        Me.mnAlmValoracionInve.Enabled = False
        
        'Facturacion Ventas
        Me.mnFacFacturarAlb.Enabled = False
        Me.mnFacContFactu.Enabled = False
        
        'Facturacion Compras
        Me.mnComFacturar.Enabled = False
        Me.mnComContFactu.Enabled = False
        
        'Reparaciones
        Me.mnRepFactAlb.Enabled = False
        
        'Mantenimientos
       ' Me.mnManFactAlb.Enabled = False
    End If
End Sub



Private Sub LanzaHome(Opcion As String)
Dim i As Integer
Dim Cad As String

    On Error GoTo ELanzaHome

'    LanzaHome = False
    'Obtenemos la pagina web de los parametros
    CadenaDesdeOtroForm = DevuelveDesdeBDNew(conAri, "spara1", Opcion, "codigo", "1", "N")
    If CadenaDesdeOtroForm = "" Then
        MsgBox "Falta configurar los datos en Parámetros de la Aplicación.", vbExclamation
        Exit Sub
    End If

    If Opcion = "webversion" Then CadenaDesdeOtroForm = CadenaDesdeOtroForm & "?version=" & App.Major & "." & App.Minor & "." & App.Revision


'    I = FreeFile
'    cad = ""
'    Open App.Path & "\lanzaexp.dat" For Input As #I
'    Line Input #I, cad
'    Close #I

    'Lanzamos
    If LanzaHomeGnral(CadenaDesdeOtroForm) Then Espera 2
    
'    If cad <> "" Then Shell cad & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'    If vConfig.Explorador <> "" Then
'        Shell vConfig.Explorador & " " & CadenaDesdeOtroForm, vbMaximizedFocus
'        LanzaHome = True
'    End If
ELanzaHome:
    If Err.Number <> 0 Then MuestraError Err.Number, Cad & vbCrLf & Err.Description
    CadenaDesdeOtroForm = ""
End Sub



Private Sub LeerEditorMenus()
Dim Sql As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    TieneEditorDeMenus = False
    Sql = "Select count(*) from usuarios.appmenus where aplicacion='Aritaxi'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        If Not IsNull(miRsAux.Fields(0)) Then
            If miRsAux.Fields(0) > 0 Then TieneEditorDeMenus = True
        End If
    End If
    miRsAux.Close
        

ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub PoneMenusDelEditor()
Dim T As Control
Dim Sql As String
Dim C As String
Dim miRsAux As ADODB.Recordset

    On Error GoTo ELeerEditorMenus
    
    Sql = "Select * from usuarios.appmenususuario where aplicacion='Aritaxi' and codusu = " & Val(Right(CStr(vUsu.Codigo), 3))
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""

    While Not miRsAux.EOF
        If Not IsNull(miRsAux.Fields(3)) Then
            Sql = Sql & miRsAux.Fields(3)
            If Right(miRsAux.Fields(3), 1) <> "|" Then Sql = Sql & "|"
            Sql = Sql & "·"
        End If
        miRsAux.MoveNext
    Wend
    miRsAux.Close
        
   
    If Sql <> "" Then
        Sql = "·" & Sql
        For Each T In Me.Controls
            If TypeOf T Is Menu Then
                C = DevuelveCadenaMenu(T)
                C = "·" & C & "·"
                Debug.Print C
                If InStr(1, Sql, C) > 0 Then
                    
                    'Stop
                    T.visible = False
                End If
           
            End If
        Next
    End If
ELeerEditorMenus:
    Set miRsAux = Nothing
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function DevuelveCadenaMenu(ByRef T As Control) As String

On Error GoTo EDevuelveCadenaMenu
    DevuelveCadenaMenu = T.Name & "|"
    DevuelveCadenaMenu = DevuelveCadenaMenu & T.Index & "|"
    Exit Function
EDevuelveCadenaMenu:
    Err.Clear
    
End Function



Private Sub PoneBarraMenus()
'Para cada boton de la toolbar comprobar que el menu con el que se corresponde
'esta visible y activado, y ponerle los mismos valore que tenga el menu
Dim Activado As Boolean

    On Error GoTo 0
    
    '-----------------------------------------------------------
    'Articulos
    Me.Toolbar1.Buttons(1).visible = ComprobarBotonMenuVisible(Me.mnAlmArticulos, Activado)
    Me.Toolbar1.Buttons(1).Enabled = Activado

    'Movimientos de Articulos
    Me.Toolbar1.Buttons(2).visible = ComprobarBotonMenuVisible(Me.mnAlmMovimArticulos, Activado)
    Me.Toolbar1.Buttons(2).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Clientes
    Me.Toolbar1.Buttons(5).visible = ComprobarBotonMenuVisible(Me.mnFacClientes, Activado)
    Me.Toolbar1.Buttons(5).Enabled = Activado
    
    'Proveedores
    Me.Toolbar1.Buttons(6).visible = ComprobarBotonMenuVisible(Me.mnComProveedores, Activado)
    Me.Toolbar1.Buttons(6).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Ofertas Clientes
'    Me.Toolbar1.Buttons(9).visible = ComprobarBotonMenuVisible(Me.mnFacOfertas(0), Activado)
'    Me.Toolbar1.Buttons(9).Enabled = Activado
    
    'Pedidos Clientes
    Me.Toolbar1.Buttons(13).visible = False
    Me.Toolbar1.Buttons(13).Enabled = Activado
    
    'Albaranes Clientes
    Me.Toolbar1.Buttons(14).visible = ComprobarBotonMenuVisible(Me.mnFacEntAlbaran, Activado)
    Me.Toolbar1.Buttons(14).Enabled = Activado
    
    'Facturas Clientes
    Me.Toolbar1.Buttons(15).visible = ComprobarBotonMenuVisible(Me.mnFacHcoFacturas, Activado)
    Me.Toolbar1.Buttons(15).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Pedidos Proveedor
    'Comprobar que los menus del que cuelga no este bloqueado o invisible
        Me.Toolbar1.Buttons(18).visible = ComprobarBotonMenuVisible(Me.mnComPedMant, Activado)
        Me.Toolbar1.Buttons(18).Enabled = Activado
    
    'Albaranes Proveedor
    Me.Toolbar1.Buttons(19).visible = ComprobarBotonMenuVisible(Me.mnComAlbMan, Activado)
    Me.Toolbar1.Buttons(19).Enabled = Activado
    
    'Facturas Proveedor
    Me.Toolbar1.Buttons(20).visible = ComprobarBotonMenuVisible(Me.mnComHcoFacturas, Activado)
    Me.Toolbar1.Buttons(20).Enabled = Activado
    
    'Recepcion facturas de compras
    Me.Toolbar1.Buttons(21).visible = ComprobarBotonMenuVisible(Me.mnComFacturar, Activado)
    Me.Toolbar1.Buttons(21).Enabled = Activado


    '-----------------------------------------------------------
    'Mantenimientos
   ' Me.Toolbar1.Buttons(21).visible = ComprobarBotonMenuVisible(Me.mnManEntrada, Activado)
   ' Me.Toolbar1.Buttons(21).Enabled = Activado
    
    'Nº Serie
    Me.Toolbar1.Buttons(25).visible = ComprobarBotonMenuVisible(Me.mnRepNumSerie, Activado)
    Me.Toolbar1.Buttons(25).Enabled = Activado
    
    
    '-----------------------------------------------------------
    'Conuslta de precio
    Me.Toolbar1.Buttons(27).visible = False
    Me.Toolbar1.Buttons(27).Enabled = Activado
      
    
    'Nuevos botones
    'TPV
'    Me.Toolbar1.Buttons(26).visible = ComprobarBotonMenuVisible(mnTPVpantallaVenta, Activado)
    
    'Cambio empresa
    Me.Toolbar1.Buttons(30).visible = ComprobarBotonMenuVisible(mnCambioEmpresa, Activado)
    Me.Toolbar1.Buttons(30).Enabled = Activado
    
'   Me.Toolbar1.Buttons(28).Image = 24 'FRECUENCIAS
    
End Sub




Private Function ComprobarBotonMenuVisible(objMenu As Menu, Activado As Boolean) As Boolean
'Comprueba si el icono de la barra se debe activar/desactivar o si se debe poner
'visible o invisible. Para ello comprueba si su correspondiente entrada de menu
'esta activada/desactiva o visible/invisible
'(se comprueba hasta q se encuentra el false o se llega al padre)
Dim nomMenu As String
Dim Sql As String
Dim RS As ADODB.Recordset
Dim Cad As String
Dim b As Boolean


    On Error GoTo EComprobar
    
    b = objMenu.visible
    Activado = objMenu.Enabled
    
    If b = False Then
        ComprobarBotonMenuVisible = b
    Else
    
        nomMenu = objMenu.Name
        
        Set RS = New ADODB.Recordset
        
        'Obtener el padre del menu
        Sql = "select padre from usuarios.appmenus where aplicacion='Aritaxi' and name=" & DBSet(nomMenu, "T")
        RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            Cad = RS.Fields(0).Value
        End If
        RS.Close
        
        b = True
        While b And Cad <> ""
                Sql = "Select name,padre from usuarios.appmenus where aplicacion='Aritaxi' and contador= " & Cad
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    Cad = RS!Padre
                    nomMenu = RS!Name
                End If
                RS.Close
                
                'comprobar si el padre esta bloqueado
                Sql = "Select count(*) from usuarios.appmenususuario where aplicacion='Aritaxi' and codusu=" & Val(Right(CStr(vUsu.Codigo), 3))
                Sql = Sql & " and tag='" & nomMenu & "|'"
                RS.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                If RS.Fields(0).Value > 0 Then
                    'Esta bloqueado el menu para el usuario
                    b = False
                End If
                RS.Close
                If Cad = "0" Then Cad = "" 'terminar si llegamos a la raiz
        Wend
        ComprobarBotonMenuVisible = b
        Set RS = Nothing
    End If
    
EComprobar:
    If Err.Number <> 0 Then Err.Clear
End Function



Private Sub AbrirListado2(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado2.Opcion = KOpcion
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub









