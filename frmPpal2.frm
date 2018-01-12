VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmPpal2 
   BackColor       =   &H00858585&
   Caption         =   "Aritaxi"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   135
   ClientWidth     =   12780
   Icon            =   "frmPpal2.frx":0000
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
            Picture         =   "frmPpal2.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7264
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":8688
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":909A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":9AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":A4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":AED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":B8E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":C2F4
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
            Picture         =   "frmPpal2.frx":CD06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":DD98
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EE2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":FEBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":10F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":129D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":13A62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":14AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":15B86
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":16C18
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":17CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":18D3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":19DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":1AE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":1BEF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":1CF84
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":1E016
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":1F0A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":2013A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":21ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":2832E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":2C830
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":2D242
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":30634
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":36E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":3D6F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":3E78A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":3F81C
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":408AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":41940
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":429D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":49234
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":4A2C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":50B28
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":51BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":52C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":53CDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":54D70
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   0
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
            Picture         =   "frmPpal2.frx":55E02
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
            TextSave        =   "11:20"
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
            Picture         =   "frmPpal2.frx":593C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":5B0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":61374
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":61D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":62798
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":64F4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":65824
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":660FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":669D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":672B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":67CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":6811E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":68230
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":68342
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":68454
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":6876E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":6E390
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":6EDA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":6F7B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":6F8C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":702D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":70CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":716FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":71A16
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":71D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":72182
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":725D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":72A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":72E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":732CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7371C
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":73A36
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":73B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":73EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":741C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":74A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":75378
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":75692
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":757EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":75B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":76518
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":76F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7793C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7834E
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":78D60
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
            Picture         =   "frmPpal2.frx":79772
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7B104
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7CA96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7E428
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":7FDBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":8174C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":830DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":84A70
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":8B2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":90AC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListMAIL 
      Left            =   360
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
            Picture         =   "frmPpal2.frx":97326
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":97778
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":97BCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":9801C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":9846E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":988C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":98D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":99164
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":995B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":9F850
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":A0262
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":A64FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":ACD5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":B35C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":B9E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":C0684
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":C6EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CD748
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CDB9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CDFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CE43E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CE890
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CECE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":CF134
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":D4D56
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":D5BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":D5EC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":D61DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":D64F6
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
      Top             =   3030
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
            Picture         =   "frmPpal2.frx":D6810
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":D851A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":DE7C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":DF1D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":DFBE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E2396
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E2C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E354A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E3E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E46FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E5110
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E556A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E567C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E578E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E58A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":E5BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EB7DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EC1EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":ECC00
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":ECD12
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":ED724
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EE136
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EEB48
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EEE62
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EF17C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EF5CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EFA20
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":EFE72
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F02C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F0716
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F0B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F0E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F0FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F12F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F1610
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F1EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F27C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F2ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F2C38
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F2F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F3964
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F4376
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F4D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F579A
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F61AC
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
            Picture         =   "frmPpal2.frx":F6BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F75D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F766B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal2.frx":F807D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
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
Attribute VB_Name = "frmPpal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Dim ContextEvent As CalendarEvent


Dim MRUShortcutBarWidth


Const IMAGEBASE = 10000
Const MinimizedShortcutBarWidth = 32 + 8

Dim WithEvents statusBar  As XtremeCommandBars.statusBar
Attribute statusBar.VB_VarHelpID = -1
Dim FontSizes(4) As Integer
Dim RibbonSeHaCreado As Boolean
Dim Pane As Pane
Dim Cad As String

'Variables comunes para todos los procedimientos de carga menus en el ribbon
'Codejock
Dim TabNuevo As RibbonTab
Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup
Dim GroupManageCalendars As RibbonGroup, GroupShare As RibbonGroup, GroupFind As RibbonGroup

Dim Control As CommandBarControl
Dim ControlNew_NewItems As CommandBarPopup
Dim Rn2 As ADODB.Recordset
Dim Habilitado As Boolean



Dim PrimeraVez As Boolean

Dim TieneEditorDeMenus As Boolean


Sub LoadResources(DllName As String, IniFileName As String)
Dim elpath As String
    
    elpath = App.Path & "\Styles\"
    CommandBarsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ShortcutBarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    SuiteControlsGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    CalendarGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    ReportControlGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
    DockingPaneGlobalSettings.ResourceImages.LoadFromFile elpath & DllName, IniFileName
End Sub




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
    Dim I As Integer
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

'    frmIdentifica.pLabel "Carga DLL"
'    CargaIconosDlls
'
'    CommandBarsGlobalSettings.App = App
'
'    frmIdentifica.pLabel "Leyendo menus usuario"
'    CargaDatosMenusDemas
'
'    ShowEventInPane = False
'
'    FontSizes(0) = 0
'    FontSizes(1) = 11
'    FontSizes(2) = 13
'    FontSizes(3) = 16
'
'    DockingPaneManager.SetCommandBars Me.CommandBars
'
'    Set frmShortBar = New frmShortcutBar2
'    Set frmInbox = New frmInbox
'
'    Dim A As Pane, b As Pane, C As Pane, d As Pane
'
'    frmIdentifica.pLabel "Creando paneles"
'    Set A = DockingPaneManager.CreatePane(PANE_SHORTCUTBAR, 170, 120, DockLeftOf, Nothing)
'    A.Tag = PANE_SHORTCUTBAR
'    A.MinTrackSize.Width = MinimizedShortcutBarWidth
'
'    Set b = DockingPaneManager.CreatePane(PANE_REPORT_CONTROL, 700, 400, DockRightOf, A)
'    b.Tag = PANE_REPORT_CONTROL
'
'    DockingPaneManager.Options.HideClient = True
'    PonerTabPorDefecto -1
'
'    Set CommandBars.Icons = CommandBarsGlobalSettings.Icons
'    LoadIcons
'
'    DockingPaneManager.RecalcLayout
'    MRUShortcutBarWidth = frmShortBar.ScaleWidth
'
'
'    'En funcion
'    ' ID_OPTIONS_STYLEBLUE2010  ID_OPTIONS_STYLESILVER2010    ID_OPTIONS_STYLEBLACK2010
'    frmIdentifica.pLabel "Carga skin"
'    Screen.MousePointer = vbHourglass
'    If vUsu.Skin = 3 Then
'        Cad = ID_OPTIONS_STYLEBLACK2010
'    Else
'        If vUsu.Skin = 2 Then
'            Cad = ID_OPTIONS_STYLESILVER2010
'        Else
'            Cad = ID_OPTIONS_STYLEBLUE2010
'        End If
'    End If
'    CommandBars.FindControl(, Cad, , True).Execute


    PrimeraVez = True
    
    
    
'--quitado
'    'Botones
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
Dim I As Integer

    CadenaDesdeOtroForm = OtrosPCsContraContabiliad
    If CadenaDesdeOtroForm <> "" Then
        I = 1
        Me.Tag = "Los siguientes PC's están conectados a: " & vEmpresa.nomempre & " (" & vUsu.CadenaConexion & ")" & vbCrLf & vbCrLf
        Do
            Sql = RecuperaValor(CadenaDesdeOtroForm, I)
            If Sql <> "" Then Me.Tag = Me.Tag & "    - " & Sql & vbCrLf
            I = I + 1
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
'--quitado
'    Me.Toolbar1.Enabled = Habilitar
'    Me.Toolbar1.visible = Habilitar
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
Dim I As Integer
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






Public Sub CambiarEmpresa(QueEmpresa As Integer)
Dim cur As Integer
    cur = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    Me.Hide
'    CambiarEmpresa2 QueEmpresa
    Me.Show
'--quitado
'    If vParam.SIITiene Then
'        DoEvents
'        Screen.MousePointer = vbHourglass
'        Espera 2.25
'        AbrirFormSII_2 False
'    End If
    Screen.MousePointer = cur
    
End Sub



'Public Sub CambiarEmpresa2(QueEmpresa As Integer)
'Dim RB As RibbonBar
'    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
'
'
'    Set vUsu = New Usuario
'    vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
'
'    vUsu.CadenaConexion = "aritaxi" & QueEmpresa
'
''    vUsu.LeerFiltros "ariconta", 301 ' asientos
''    vUsu.LeerFiltros "ariconta", 401 ' facturas de cliente
'
'    AbrirConexion
'
'    Set vEmpresa = New Cempresa
'    Set vParam = New Cparametros
'    'NO DEBERIAN DAR ERROR
'    vEmpresa.LeerDatos
'    vParam.Leer
'
'    PonerCaption
'
'    Screen.MousePointer = vbHourglass
'
'   CargaDatosMenusDemas
'   frmPaneContacts.SeleccionarNodoEmpresa vEmpresa.codempre
'   pageBackstageHelp.Label9.Caption = vEmpresa.nomempre
'   pageBackstageHelp.tabPage(0).visible = False
'   pageBackstageHelp.tabPage(1).visible = False
'   frmInbox.OpenProvider
'   Set RB = RibbonBar
'   RB.Minimized = False
'   RB.RedrawBar
'
'
'
'
'   Screen.MousePointer = vbDefault
'End Sub

Private Sub PonerCaption()
        Caption = "AriTAXI 6    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.Nombre
        'Label33.Caption = "   " & vEmpresa.nomempre
End Sub


'Private Sub CargaDatosMenusDemas()
'Dim AntiguoTab As Integer
'
'
'    Screen.MousePointer = vbHourglass
'    AntiguoTab = -1
'    If RibbonSeHaCreado Then
'        If Not RibbonBar.SelectedTab Is Nothing Then AntiguoTab = RibbonBar.SelectedTab.Id
'    End If
'    CreateRibbon
'    Screen.MousePointer = vbHourglass
'    CreateBackstage
'    Screen.MousePointer = vbHourglass
'    CreateRibbonOptions
'
'    vEmpresa.LeerDatos
'
'    'vEmpresa.TieneContabilidad = False
'    '??????
'    '0=solo contabilidad / 1=todo / 2=solo tesoreria
'    Screen.MousePointer = vbHourglass
'    CargaMenu AntiguoTab
'    CreateStatusBar
'    Screen.MousePointer = vbHourglass
'    PonerCaption
'    CreateCalendarTabOriginal
'    RibbonSeHaCreado = True
'End Sub
'


'*************************************************************************
'*************************************************************************
'*************************************************************************
'
'       CARGA menus en Ribbon
'
'




'Public Sub CargaMenu(AntiguoTab As Integer)
'Dim RN As ADODB.Recordset
'
'
'
'
'    Set RN = New ADODB.Recordset
'    Set Rn2 = New ADODB.Recordset
'    On Error GoTo eCargaMenu
'
'
'    If RibbonSeHaCreado Then RibbonBar.RemoveAllTabs
'
'    Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =0 ORDER BY padre,orden "
'    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not RN.EOF
'
'
'        If Not BloqueaPuntoMenu(RN!Codigo, "aritaxi") Then
'             Habilitado = True
'
'             If Not MenuVisibleUsuario(DBLet(RN!Codigo), "aritaxi") Then
'                 Habilitado = False
'             Else
'
'                 If (MenuVisibleUsuario(DBLet(RN!Padre), "aritaxi") And DBLet(RN!Padre) <> 0) Or DBLet(RN!Padre) = 0 Then
'                     'OK todo habilitado
'                 Else
'                     Habilitado = False
'                 End If
'             End If
'
'
'
'            If Habilitado Then
'
'                Select Case RN!Codigo
'                Case 1
'                    '1   "CONFIGURACION"
'                    CargaMenuConfiguracion RN!Codigo
''                Case 2
''                    '2   "DATOS GENERALES"
''                    CargaMenuDatosGenerales RN!Codigo
''                Case 3
''                    '3   "DIARIO"
''                    CargaMenuDiarios RN!Codigo
''                Case 4
''                    '4   "FACTURAS"
''                    CargaMenuFacturas RN!Codigo
''                Case 5
''                    '5   "INMOVILIZADO"
''                    CargaMenuInmovilizado RN!Codigo
''                Case 6
''                    '6   "CARTERA DE COBROS"
''                    CargaMenuTesoreriaCobros RN!Codigo
''                Case 7
''
''                Case 8
''                    '8   "CARTERA DE PAGOS"
''                    CargaMenuTesoreriaPagos RN!Codigo
''                Case 9
''                    '9   "INFORMES TESORERIA"
''                     CargaMenuTesoreriaInformes RN!Codigo
''                Case 10
''                    '10  "ANALÍTICA"
''                    'Va dentro de diario
''                    'UNa solapa para el
''                    CargaMenuAnaliticaPResupuestaria RN!Codigo
''                Case 11
''                    '11  "PRESUPUESTARIA"
''                    CargaMenuAnaliticaPResupuestaria RN!Codigo
''                Case 12
''
''                Case 13
''                    '13  "CIERRE EJERCICIO"
''                    CargaMenuCierreEjercicio RN!Codigo
''
''                Case 14
''                    '14  "UTILIDADES"
''                    CargaMenuUtilidades RN!Codigo
''                Case Else
''                    MsgBox "Menu no tratado"
''                    End
'                End Select
'
'            End If
'
'        End If  'de habilitado el padre
'
'        RN.MoveNext
'    Wend
'    RN.Close
'
'    PonerTabPorDefecto AntiguoTab
'
'eCargaMenu:
'    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
'
'    Set TabNuevo = Nothing
'    Set GroupNew = Nothing
'    Set Control = Nothing
'    Set RN = Nothing
'    Set Rn2 = Nothing
'End Sub
'

Public Function BloqueaPuntoMenu(IdProg As Long, Aplicacion As String) As Boolean
Dim EsdeAnalitica As Boolean

    BloqueaPuntoMenu = False

    If Aplicacion = "aritaxi" Then
        ' programas de analitica
'--revisar
'        EsdeAnalitica = (IdProg = 10 Or IdProg = 1001 Or IdProg = 1002 Or IdProg = 1003 Or IdProg = 1004 Or IdProg = 1005)
'        BloqueaPuntoMenu = (Not vParam.Autocoste And EsdeAnalitica)
    End If
    
End Function



Public Function MenuVisibleUsuario(Proceso As Long, Aplicacion As String) As Boolean
Dim Sql As String
Dim Excepcion As String


    Sql = "select ver from menus_usuarios where codigo = " & DBSet(Proceso, "N") & " and aplicacion = " & DBSet(Aplicacion, "T")
    Sql = Sql & " and codusu = " & DBSet(vUsu.Id, "N")
    
    MenuVisibleUsuario = (DevuelveValor(Sql) = 1)

End Function



'Private Sub PonerTabPorDefecto(AntiguoTabSeleccionado As Integer)
'Dim Anterior As Integer
'
'    On Error Resume Next
'
'    If AntiguoTabSeleccionado < 0 Then
'        Anterior = vUsu.TabPorDefecto
'    Else
'        Anterior = AntiguoTabSeleccionado
'    End If
'
'    Cad = ""
'    For i = 0 To RibbonBar.TabCount - 1
'        J = RibbonBar.Tab(i).Id
'        'Debug.Print J & " " & RibbonBar.Tab(i).Caption
'        If J = Anterior Then
'
'            RibbonBar.Tab(i).visible = True
'            RibbonBar.Tab(i).Selected = True
'            Set RibbonBar.SelectedTab = RibbonBar.Tab(i)
'            Cad = "OK"
'            Exit For
'        End If
'    Next
'    If Cad = "" Then
'
'        For J = RibbonBar.TabCount To 1 Step -1
'            RibbonBar.Tab(J - 1).visible = True
'            RibbonBar.Tab(J - 1).Selected = True
'        Next J
'    End If
'
'    Err.Clear
'End Sub

'Private Sub CargaMenuConfiguracion(IdMenu As Integer)
'
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Configuracion")
'        TabNuevo.Id = CLng(IdMenu)
'        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
'
'
'
'        'todos los hijos que cuelgan en la tab
'        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        Cad = ""
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
'                End If
'
'
'
'                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'                Control.Enabled = Habilitado
'
'            End If
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'        'color Categorias  eventos
'        If Not GroupNew Is Nothing Then
'            Set Control = GroupNew.Add(xtpControlButton, 199, "Categorias calendario")
'        End If
'        Set GroupNew = Nothing
'End Sub
'
'
'
'
'
'
'Private Sub CargaMenuDatosGenerales(IdMenu As Integer)
'Dim SegundoGrupo As RibbonGroup
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Datos generales")
'        TabNuevo.Id = CLng(IdMenu)
'
'
'        'En este llevaremos dos solapas, tesoreria y contabilidad (no le ponemos nombres)
'        Cad = CStr(IdMenu * 100000)
'
''--quitado
''        If vEmpresa.TieneContabilidad Then Set GroupNew = TabNuevo.Groups.AddGroup("", cad & "0")
''        If vEmpresa.TieneTesoreria Then Set SegundoGrupo = TabNuevo.Groups.AddGroup("", cad & "1")
'
'        Set GroupNew = TabNuevo.Groups.AddGroup("", Cad & "0")
'
'
'        'todos los hijos que cuelgan en la tab
'        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        Cad = ""
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
'                End If
'
'
'
'                If Rn2!Tipo = 1 Then
'                    Set Control = SegundoGrupo.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'                Else
'                    Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'                End If
'
'                Control.Enabled = Habilitado
'                ' Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
'                '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Appointment")
'                '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "All Day E&vent")
'                '     Control.BeginGroup = True
'                ' ControlNew_NewItems.KeyboardTip = "V"
'
'            End If
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'         Set GroupNew = Nothing
'End Sub
'

Private Sub CargaMenuDiarios(IdMenu As Integer)
Dim GrupSald As RibbonGroup
Dim GrOtro As RibbonGroup
Dim GrConsoli As RibbonGroup

'        If Not vEmpresa.TieneContabilidad Then Exit Sub
'
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Diario")
'        TabNuevo.id = CLng(IdMenu)
'
'        cad = CStr(IdMenu * 100000)
'        Set GroupNew = TabNuevo.Groups.AddGroup("ASIENTOS", cad & "0")
'        Set GrupSald = TabNuevo.Groups.AddGroup("BALANCES", cad & "1")
'        Set GrOtro = TabNuevo.Groups.AddGroup("", cad & "2")
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'
'
'
'                Select Case Rn2!Codigo
'                Case 301, 303, 304, 314, 211
'                    Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'                Case 306, 307, 308, 309
'                    Set Control = GrupSald.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'                'Consolidado
'                Case 315
'                    Set GrConsoli = TabNuevo.Groups.AddGroup("CONSOLIDADO", cad & "4")
'                    Set Control = GrConsoli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'                Case Else
'                    Set Control = GrOtro.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'                End Select
'
'
'                Control.Enabled = Habilitado
'
'
'
'
'            End If
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'    Set GrupSald = Nothing
'    Set GrOtro = Nothing
'     Set GrConsoli = Nothing
End Sub


Private Sub CargaMenuFacturas(IdMenu As Integer)
'Dim GropCli As RibbonGroup
'Dim GrupPag As RibbonGroup
'Dim Consoli As RibbonGroup
'Dim OpsAseg As RibbonGroup
'Dim Insertado As Boolean
'Dim B As Boolean
'
''        If Not vEmpresa.TieneContabilidad Then Exit Sub
'
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Facturas")
'        TabNuevo.id = CLng(IdMenu)
'
'        cad = CStr(IdMenu * 100000)
'        Set GropCli = TabNuevo.Groups.AddGroup("EMITIDAS", cad & "0")
'        Set GrupPag = TabNuevo.Groups.AddGroup("RECIBIDAS", cad & "1")
'        Set GroupNew = TabNuevo.Groups.AddGroup("I.V.A.", cad & "2")
'
'
''
''        401 "Facturas Emitidas" 14
''        402 "Libro Facturas Emitidas"   16
''        403 "Relación Clientes por cuenta"  0
''        404 "Facturas Recibidas"    15
''        405 "Libro Facturas Recibidas"  17
''        406 "Relacion Proveedores por cuenta"   0
''        408 "Modelo 303"    0
''        409 "Modelo 340"    0
''        410 "Modelo 347"    0
''        411 "Modelo 349"    0
''        412 "Liquidacion I.V.A."    18
''        413 consolidado
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'            Insertado = True
'            Select Case Rn2!Codigo
'            Case 401, 402, 403
'                Set Control = GropCli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'            Case 404, 405, 406
'                Set Control = GrupPag.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            Case 413
'                Set Consoli = TabNuevo.Groups.AddGroup("CONSOLIDADO", CStr(IdMenu * 100000) & "2")
'                Set Control = Consoli.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'
'            Case 414, 415
'                 If vParamT.TieneOperacionesAseguradas Then
'                        If OpsAseg Is Nothing Then Set OpsAseg = TabNuevo.Groups.AddGroup("OP. ASEGURADAS", CStr(IdMenu * 100000) & "4")
'                        Set Control = OpsAseg.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'                 Else
'                    Insertado = False
'                 End If
'
'            Case Else
'                B = True
'                If Rn2!Codigo = ID_SII Then
'
'                    If vParam.SIITiene Then
'                        If vUsu.Nivel > 0 Then B = False
'                    Else
'                        B = False
'
'                    End If
'                End If
'                If Not B Then Habilitado = False
'
'                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            End Select
'
'
'            cad = "NO"
'            If Insertado Then Control.Enabled = Habilitado
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'
'
'
'
End Sub



Private Sub CargaMenuInmovilizado(IdMenu As Integer)
'
'        If Not vEmpresa.TieneContabilidad Then Exit Sub
'
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Inmovilizado")
'        TabNuevo.id = CLng(IdMenu)
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'            If cad = "" Then Set GroupNew = TabNuevo.Groups.AddGroup("", CStr(IdMenu * 100000) & "0")
'            cad = "NO"
'            'Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&New Appointment")
'            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'            Control.Enabled = Habilitado
'
'           ' Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
'           '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Appointment")
'           '     Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "All Day E&vent")
'           '     Control.BeginGroup = True
'           ' ControlNew_NewItems.KeyboardTip = "V"
'
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'

End Sub




Private Sub CargaMenuTesoreriaCobros(IdMenu As Integer)
'Dim GrupCob As RibbonGroup
'Dim GrupRem As RibbonGroup
'
''    601 "Cartera de Cobros"
''    602 "Informe Cobros Pendientes"
''    604 "Realizar Cobro"
''    606 "Compensaciones"
''    607 "Compensar cliente"
''    608 "Reclamaciones"
''    609 "Remesas"
''    610 "Informe Impagados"
''    611 "Recepción Talón-Pagaré"
''    612 "Remesas Talón-Pagaré"
''    613 "Norma 57 - Pagos ventanilla"
''    614 "Transferencias Abonos"
'
'
'        If Not vEmpresa.TieneTesoreria Then Exit Sub
'
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Tesoreria")
'        TabNuevo.id = CLng(IdMenu)
'        NumRegElim = TabNuevo.Index
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'
'
'        'Creamos los tres grupos
'        cad = CStr(IdMenu * 100000)
'        Set GrupCob = TabNuevo.Groups.AddGroup("COBROS", cad & "0")
'        Set GrupRem = TabNuevo.Groups.AddGroup("REMESAS", cad & "1")
'        Set GroupNew = TabNuevo.Groups.AddGroup("", cad & "2")
'
'
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'
'
'            Select Case Rn2!Codigo
'            Case 601, 602, 604, 607, 608, 610, 613, 614
'                'Solapa cobros
'                Set Control = GrupCob.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            Case 609, 611, 612
'                'Solapa remesas
'                Set Control = GrupRem.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            Case Else
'                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'
'
'            End Select
'
'
'            Control.Enabled = Habilitado
'
'           ' ControlNew_NewItems.KeyboardTip = "V"
'
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'
End Sub



Private Sub CargaMenuTesoreriaPagos(IdMenu As Integer)
'
'
'        If Not vEmpresa.TieneTesoreria Then Exit Sub
'
'
'
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'
'
'        'Los pagos se gargan sobre la solapa de TESORERIA
'
'        Set GroupNew = RibbonBar.Tab(NumRegElim).Groups.AddGroup("PAGOS", CStr(IdMenu * 100000) & "0")
'
'
'        While Not Rn2.EOF
'
''            801 "Cartera de Pagos"  5
''            802 "Informe Pagos pendientes"  19
''            803 "Informe Pagos bancos"  0
''            804 "Realizar Pago" 24
''            805 "Transferencias"    0
''            806 "Pagos domiciliados"    0
''            807 "Gastos Fijos"  0
''            809 "Compensar proveedor"   0
''            810 "Confirming"    0
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            Control.Enabled = Habilitado
'
'
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'

End Sub


Private Sub CargaMenuTesoreriaInformes(IdMenu As Integer)
'
'
'        If Not vEmpresa.TieneTesoreria Then Exit Sub
'
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu
'        'De momento NO cargamos el 904
'        cad = cad & " AND codigo <>904  ORDER BY padre,orden"
'
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'
'
'        'Los informes se gargan sobre la solapa de TESORERIA
'
'        Set GroupNew = RibbonBar.Tab(NumRegElim).Groups.AddGroup("INFORMES", CStr(IdMenu * 100000) & "0")
'
'
'        While Not Rn2.EOF
''       901 "ariconta"  9   "Informe por NIF *" 1   1   0
''       902 "ariconta"  9   "Informe por cuenta *"  2   1   0
''       903 "ariconta"  9   "Situación Tesoreria *" 3   1   29
''       904 "ariconta"  9   "Memoria Plazos de pago *"  4   1   0
''     ID_InformeporNIF ID_Informeporcuenta ID_SituaciónTesoreria
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            Control.Enabled = Habilitado
'
'
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'

End Sub




Private Sub CargaMenuAnaliticaPResupuestaria(IdMenu As Integer)
'
'
'
'        If Not vEmpresa.TieneContabilidad Then Exit Sub
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'
'
'        'Los pagos se gargan sobre la solapa de diario
'        Set TabNuevo = RibbonBar.FindTab(3)
'        If TabNuevo Is Nothing Then
'            Set TabNuevo = RibbonBar.FindTab(15)
'            If TabNuevo Is Nothing Then
'
'                Set TabNuevo = RibbonBar.InsertTab(15, "ANALITICA-PRESUPUESTO")
'                TabNuevo.id = 15
'
'            End If
'        End If
'        cad = CStr(IdMenu * 100000) & "0"
'        Set GroupNew = TabNuevo.Groups.AddGroup(IIf(IdMenu = 10, "ANALITICA", "PRESUPUESTOS"), cad)
'
'
'        While Not Rn2.EOF
'
''            801 "Cartera de Pagos"  5
''            802 "Informe Pagos pendientes"  19
''            803 "Informe Pagos bancos"  0
''            804 "Realizar Pago" 24
''            805 "Transferencias"    0
''            806 "Pagos domiciliados"    0
''            807 "Gastos Fijos"  0
''            809 "Compensar proveedor"   0
''            810 "Confirming"    0
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'
'
'
'            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'
'            Control.Enabled = Habilitado
'
'           ' ControlNew_NewItems.KeyboardTip = "V"
'
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'
End Sub




Private Sub CargaMenuCierreEjercicio(IdMenu As Integer)
'Dim GropCli As RibbonGroup
'Dim GrupPag As RibbonGroup
'
'        If Not vEmpresa.TieneContabilidad Then Exit Sub
'
'        'Creamos la TAB
'        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Cierre ejercicio")
'        TabNuevo.id = CLng(IdMenu)
'
'        Set GroupNew = TabNuevo.Groups.AddGroup("", 13000001)
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'        While Not Rn2.EOF
'
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
''        1301    "Renumeración de asientos"  0
''        1303    "Cierre de Ejercicio"   0
''        1304    "Deshacer cierre"   0
''        1306    "Diario Oficial"    0
''        1308    "Presentación Telemática de Libros" 0
''        1309    "Memoria Plazos de pago"    0
'
'
'
'            Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
'            Control.Enabled = Habilitado
'            ' ControlNew_NewItems.KeyboardTip = "V"
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'
End Sub

Private Function DevulevePosicionUtilidades(Id As Integer) As Integer
    Select Case Id
    Case ID_Traspasodecuentasenapuntes
        DevulevePosicionUtilidades = 1
    Case ID_Renumerarregistrosproveedor
        DevulevePosicionUtilidades = 2
    Case ID_Aumentardígitoscontables
        DevulevePosicionUtilidades = 3
    Case ID_TraspasocodigosdeIVA
        DevulevePosicionUtilidades = 4
    Case Else
        'ID_Accionesrealizadas
        DevulevePosicionUtilidades = 5
    End Select
End Function

Private Sub CargaMenuUtilidades(IdMenu As Integer)
'Dim Col As Collection
'
'
'
'
'        'Este veremos si tiene alguna utilidad activa. Si es asi, crearemos la solapa, si no nada
'        '.......................................................................
'
'
'        'todos los hijos que cuelgan en la tab
'        cad = "Select * from menus where aplicacion = 'ariconta' and padre =" & IdMenu & " ORDER BY padre,orden"
'        Rn2.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        cad = ""
'        Set Col = New Collection
'        While Not Rn2.EOF
'           i = i + 1
'           If Not BloqueaPuntoMenu(Rn2!Codigo, "ariconta") Then
'                Habilitado = True
'
'                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "ariconta") Then
'                    Habilitado = False
'                Else
'                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "ariconta") Then Habilitado = False
'                End If
'            End If
'
'            If Rn2!Codigo = 1414 Then
'                If vParamT.FormaPagoInterTarjeta < 0 Then Habilitado = False
'            End If
'
'            Col.Add Abs(Habilitado) & "|" & Rn2!Codigo & "|" & Rn2!Descripcion & "|"
'            If Habilitado Then cad = "S"
'
'            Rn2.MoveNext
'        Wend
'        Rn2.Close
'
'            '1408    "Traspaso de cuentas en apuntes"
'            '1409    "Renumerar registros proveedor"
'            '1410    "Aumentar dígitos contables"
'            '1411    "Traspaso códigos de I.V.A."
'            '1412    "Acciones realizadas"
'            '1413    Importar fras cliente
'            '1414    importacon facturas (de momento consum)
'
'        'Ya puedo utilizar numregelim
'        If cad <> "" Then
'            'OK creamos solapa y demas
'            'Creamos la TAB
'            Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Utilidades")
'            TabNuevo.id = CLng(IdMenu)
'            Set GroupNew = TabNuevo.Groups.AddGroup("", 14000001)
'            For NumRegElim = 1 To Col.Count
'                Habilitado = CStr(RecuperaValor(Col.Item(NumRegElim), 1)) = "1"
'                Set Control = GroupNew.Add(xtpControlButton, CLng(RecuperaValor(Col.Item(NumRegElim), 2)), CStr(RecuperaValor(Col.Item(NumRegElim), 3)))
'                Control.Enabled = Habilitado
'            Next
'
'
'        End If
'
'
'Set Col = Nothing
End Sub

'Public Sub CheckButton(nButton As Integer)
'    CommandBars.Actions(ID_OPTIONS_STYLEBLUE2010).Checked = False
'    CommandBars.Actions(ID_OPTIONS_STYLESILVER2010).Checked = False
'    CommandBars.Actions(ID_OPTIONS_STYLEBLACK2010).Checked = False
'
'    CommandBars.Actions(nButton).Checked = True
'End Sub

'
'Sub OnThemeChanged(Id As Integer)
'Dim N_Skin As Integer
'    CheckButton Id
'
'    Dim FlatStyle As Boolean
'    FlatStyle = Id >= ID_OPTIONS_STYLESCENIC7 And Id <= ID_OPTIONS_STYLEBLACK2010
'
'
'    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
'
'
'    CommandBars.EnableOffice2007Frame False
'
'    Select Case CommandBars.VisualTheme
'        Case xtpThemeResource, xtpThemeRibbon
'            CommandBars.AllowFrameTransparency False 'True
'            CommandBars.EnableOffice2007Frame True
'            CommandBars.SetAllCaps False
'            CommandBars.statusBar.SetAllCaps False
'        Case Else
'            CommandBars.AllowFrameTransparency True
'            CommandBars.EnableOffice2007Frame False
'            CommandBars.SetAllCaps False
'            CommandBars.statusBar.SetAllCaps False
'    End Select
'
'    Dim ToolTipContext As ToolTipContext
'    Set ToolTipContext = CommandBars.ToolTipContext
'    ToolTipContext.Style = xtpToolTipResource
'    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
'    ToolTipContext.ShowImage True, IMAGEBASE
'    ToolTipContext.SetMargin 2, 2, 2, 2
'    ToolTipContext.MaxTipWidth = 180
'
'    statusBar.ToolTipContext.Style = ToolTipContext.Style
'    frmShortBar.wndShortcutBar.ToolTipContext.Style = ToolTipContext.Style
'
'
'    'CreateBackstage
'    'SetBackstageTheme
'
'    'CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\styles\Ariconta.ico", 16, 16
'
'    'Set Captions VisualTheme
'    On Error Resume Next
'    Dim CtrlCaption As ShortcutCaption
'    Dim Form As Form, Ctrl As Object
'
'    For Each Form In Forms
'        For Each Ctrl In Form.Controls
'
'            Set CtrlCaption = Ctrl
'            If Not CtrlCaption Is Nothing Then
'                CtrlCaption.VisualTheme = frmShortBar.wndShortcutBar.VisualTheme
'            End If
'
'        Next
'    Next
'
'    DockingPaneManager.PaintManager.SplitterSize = 5
'    DockingPaneManager.PaintManager.SplitterColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
'
'    DockingPaneManager.PaintManager.ShowCaption = False
'    DockingPaneManager.RedrawPanes
'
'    frmShortBar.SetColor Id
'    frmInbox.SetColor Id
'
'
'    frmPaneCalendar.SetFlatStyle FlatStyle
'    frmPaneContacts.SetFlatStyle FlatStyle
'    'frmPaneInformacion.SetFlatStyle FlatStyle
'    'frmPaneAcercaDe.SetFlatStyle FlatStyle
'
'
'
'
'
'
'    LoadIcons
'    N_Skin = Id - 2895
'    EstablecerSkin N_Skin
'
'    'Updatear SKIN usuario
'    If CStr(N_Skin) <> vUsu.Skin Then
'        vUsu.Skin = N_Skin
'        vUsu.ActualizarSkin
'    End If
'
'End Sub
'
'
'Private Sub CreateRibbon()
'    Dim RibbonBar As RibbonBar
'
'    If RibbonSeHaCreado Then Exit Sub
'
'
'
'    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
'    RibbonBar.EnableDocking xtpFlagStretched
'
'    RibbonBar.AllowQuickAccessCustomization = False
'    RibbonBar.ShowQuickAccessBelowRibbon = False
'    RibbonBar.ShowGripper = False
'
'    RibbonBar.AllowMinimize = False
'    RibbonBar.AddSystemButton
'
'    RibbonBar.SystemButton.IconId = ID_SYSTEM_ICON
'    RibbonBar.SystemButton.Caption = "&Menu"
'    RibbonBar.SystemButton.Style = xtpButtonCaption
'End Sub
'
'Private Sub CreateRibbonOptions()
'
'    CommandBars.EnableActions
'    If RibbonSeHaCreado Then Exit Sub
'
'    CommandBars.Actions.Add ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue", "Office 2010 Blue", "Office 2010 Blue", "Themes"
'    CommandBars.Actions.Add ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver", "Office 2010 Silver", "Office 2010 Silver", "Themes"
'    CommandBars.Actions.Add ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black", "Office 2010 Black", "Office 2010 Black", "Themes"
'
'    Dim Control As CommandBarControl, ControlAbout As CommandBarControl
'    Dim ControlPopup As CommandBarPopup, ControlOptions As CommandBarPopup
'
'    Set ControlOptions = RibbonBar.Controls.Add(xtpControlPopup, 0, "Opciones")
'    ControlOptions.Flags = xtpFlagRightAlign
'
'    Set Control = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Styles")
'    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue"
'    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver"
'    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black"
'
'    Set ControlPopup = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Tamaño fuente", -1, False)
'    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_SYSTEM, "Sistema", -1, False
'    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlRadioButton, ID_OPTIONS_FONT_NORMAL, "Normal", -1, False)
'    Control.BeginGroup = True
'    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_LARGE, "Grande", -1, False
'    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_EXTRALARGE, "Extra grande", -1, False
'    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_OPTIONS_FONT_AUTORESIZEICONS, "Ajustar Icons", -1, False)
'    Control.BeginGroup = True
'
'    'ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_RTL, "Right To Left"
'    ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_ANIMATION, "Animation   "
'
'    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_MINIMIZE, "Minimizar la barra", False, "Muestra solo los titulos del menu principal.")
'    Control.Flags = xtpFlagRightAlign
'
'    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_EXPAND, "Expandir la barra", False, "Muestra todos los elementos del menu.")
'    Control.Flags = xtpFlagRightAlign
'
'    Set ControlAbout = RibbonBar.Controls.Add(xtpControlButton, ID_APP_ABOUT, "&Acerca de")
'    ControlAbout.Flags = xtpFlagRightAlign Or xtpFlagManualUpdate
'
'
'
'End Sub
'
'
'Private Sub CreateBackstage()
'
'
'    Dim RibbonBar As RibbonBar
'    Set RibbonBar = CommandBars.ActiveMenuBar
'
'    Dim BackstageView As RibbonBackstageView
'    Set BackstageView = CommandBars.CreateCommandBar("CXTPRibbonBackstageView")
'
'    BackstageView.SetTheme xtpThemeRibbon
'
'
'    CommandBars.Icons.LoadBitmap App.Path & "\styles\BackstageIcons.png", _
'    Array(1, 1, 1002, 1, 1, ID_APP_EXIT), xtpImageNormal
'
'    Set RibbonBar.AddSystemButton.CommandBar = BackstageView
'
'    'BackstageView.AddCommand ID_FILE_SAVE, "Cambiar empresa"
'    'BackstageView.AddCommand ID_FILE_SAVE_AS, "Personalizar"
'    'BackstageView.AddCommand ID_FILE_OPEN, "Open"
'    'BackstageView.AddCommand ID_FILE_CLOSE, "Close"
'
'    'If (pageBackstageInfo Is Nothing) Then Set pageBackstageInfo = New pageBackstageInfo
'    'If (pageBackstageSend Is Nothing) Then Set pageBackstageSend = New pageBackstageSend
'    If (pageBackstageHelp Is Nothing) Then Set pageBackstageHelp = New pageBackstageHelp
'
'    Dim ControlInfo As RibbonBackstageTab
'    Set ControlInfo = BackstageView.AddTab(1000, "Info", pageBackstageHelp.hwnd)
'
'    'BackstageView.AddTab 1002, "Empresas", pageBackstageSend.hwnd
'
'    ' Los menus de informacion...
'    'BackstageView.AddTab 1001, "Acerca de", pageBackstageInfo.hwnd
'
'
'
'
'
'
'
'
'
'
'    'BackstageView.AddCommand ID_FILE_OPTIONS, "Options"
'    BackstageView.AddCommand ID_APP_EXIT, "Salir"
'
'    ControlInfo.DefaultItem = True
'
'
'End Sub
'
'
'
'
'
'Private Sub CreateStatusBar()
'Dim Pane As StatusBarPane
'
'    If RibbonSeHaCreado Then
'        'StatusBar.Pane(0).Value = vEmpresa.nomempre & "    " & vUsu.Login
'        statusBar.Pane(0).Text = "Nº " & vEmpresa.codempre
'        statusBar.Pane(1).Text = vEmpresa.nomempre
'
'    Else
'
'
'         Set statusBar = Nothing
'
'         Set statusBar = CommandBars.statusBar
'         statusBar.visible = True
'
'
'         Set Pane = statusBar.AddPane(ID_INDICATOR_PAGENUMBER)
'         Pane.Text = "Nº " & vEmpresa.codempre
'         Pane.Caption = "&C"
'         Pane.Value = vEmpresa.nomempre & "    " & vUsu.Login
'         Pane.Button = True
'         Pane.SetPadding 8, 0, 8, 0
'
'         Set Pane = statusBar.AddPane(ID_INDICATOR_WORDCOUNT)
'         Pane.Text = vEmpresa.nomempre
'         Pane.Caption = ""
'         Pane.Value = vEmpresa.codempre
'         Pane.Button = True
'         Pane.SetPadding 8, 0, 8, 0
'
'
'         Set Pane = statusBar.AddPane(0)
'         Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
'         Pane.BeginGroup = True
'
'        '
'         statusBar.RibbonDividerIndex = 3
'         statusBar.EnableCustomization True
'
'         CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowNever
'         CommandBars.Options.ShowKeyboardTips = True
'         CommandBars.Options.ToolBarAccelTips = True
'    End If
'End Sub
'
'Private Sub DockBarRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
'    Dim Left As Long
'    Dim top As Long
'    Dim Right As Long
'    Dim Bottom As Long
'
'    CommandBars.RecalcLayout
'    BarOnLeft.GetWindowRect Left, top, Right, Bottom
'
'    CommandBars.DockToolBar BarToDock, Right, (Bottom + top) / 2, BarOnLeft.Position
'
'End Sub
'
'Private Sub CommandBars_CommandBarKeyDown(CommandBar As XtremeCommandBars.ICommandBar, KeyCode As Long, Shift As Integer)
'    Debug.Print CommandBar.BarID
'End Sub
'
'Public Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'Dim AbiertoFormulario  As Boolean
'    AbiertoFormulario = False
'
'
'    Select Case Control.Id
'        Case XTPCommandBarsSpecialCommands.XTP_ID_RIBBONCONTROLTAB:
'
'
'
'        Case XTP_ID_RIBBONCUSTOMIZE:
'            CommandBars.ShowCustomizeDialog 3
'
'        Case ID_APP_ABOUT:
'
'           LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriTAXI-6.html?"
'
'
'        Case ID_FILE_NEW:
'            'frmEmail.Show 0, Me
'
'
'
'        Case ID_Licencia_Usuario_Final_txt, ID_Licencia_Usuario_Final_web, ID_Ver_Version_operativa_web
'            OpcionesMenuInformacion Control.Id
'
'
'
'        Case ID_VIEW_STATUSBAR:
'            CommandBars.statusBar.visible = Not CommandBars.statusBar.visible
'            CommandBars.RecalcLayout
'
'        Case ID_RIBBON_EXPAND:
'            RibbonBar.Minimized = Not RibbonBar.Minimized
'
'        Case ID_RIBBON_MINIMIZE:
'            RibbonBar.Minimized = Not RibbonBar.Minimized
'
'        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
'            Dim newFontHeight As Integer
'            newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
'            RibbonBar.FontHeight = newFontHeight
'
'        Case ID_OPTIONS_FONT_AUTORESIZEICONS
'            CommandBars.PaintManager.AutoResizeIcons = Not CommandBars.PaintManager.AutoResizeIcons
'            CommandBars.RecalcLayout
'            RibbonBar.RedrawBar
'
'        Case ID_OPTIONS_STYLEBLUE2010:
'            LoadResources "Office2010.dll", "Office2010Blue.ini"
'            CommandBars.VisualTheme = xtpThemeRibbon
'            DockingPaneManager.VisualTheme = ThemeResource
'            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
'            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
'            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
'
'            OnThemeChanged ID_OPTIONS_STYLEBLUE2010
'
'
'
'       Case ID_OPTIONS_STYLESILVER2010:
'            LoadResources "Office2010.dll", "Office2010Silver.ini"
'            CommandBars.VisualTheme = xtpThemeRibbon
'            DockingPaneManager.VisualTheme = ThemeResource
'            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
'            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
'            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
'
'            OnThemeChanged ID_OPTIONS_STYLESILVER2010
'
'       Case ID_OPTIONS_STYLEBLACK2010:
'            LoadResources "Office2010.dll", "Office2010Black.ini"
'            CommandBars.VisualTheme = xtpThemeRibbon
'            DockingPaneManager.VisualTheme = ThemeResource
'            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
'            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
'            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
'
'            OnThemeChanged ID_OPTIONS_STYLEBLACK2010
'
'        Case ID_APP_EXIT:
'            Unload Me
'
'
'
'        Case ID_GROUP_GOTO_TODAY:
'            Select Case frmInbox.CalendarControl.ViewType
'                Case xtpCalendarDayView:
'                    frmInbox.CalendarControl.DayView.ShowDay DateTime.Now, True
'
'                Case xtpCalendarWorkWeekView:
'                    frmInbox.CalendarControl.DayView.SetSelection DateTime.Now, DateTime.Now, True
'                    frmInbox.CalendarControl.RedrawControl
'
'                Case xtpCalendarWeekView:
'                    frmInbox.CalendarControl.WeekView.SetSelection DateTime.Now, DateTime.Now, True
'
'                Case xtpCalendarMonthView:
'                    frmInbox.CalendarControl.MonthView.SetSelection DateTime.Now, DateTime.Now, True
'            End Select
'
'        Case ID_GROUP_GOTO_NEXT7DAYS:
'            Dim lastDate As Date
'            lastDate = frmInbox.CalendarControl.DayView.Days(frmInbox.CalendarControl.DayView.DaysCount - 1).Date
'            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
'            frmInbox.CalendarControl.DayView.ShowDays lastDate + 1, lastDate + 7
'
'        Case ID_GROUP_ARRANGE_DAY:
'            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
'
'        Case ID_GROUP_ARRANGE_WORK_WEEK:
'            frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView
'
'        Case ID_GROUP_ARRANGE_WEEK:
'            frmInbox.CalendarControl.UseMultiColumnWeekMode = True
'            frmInbox.CalendarControl.ViewType = xtpCalendarWeekView
'
'        Case ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_MONTH_LOW, _
'             ID_GROUP_ARRANGE_MONTH_MEDIUM, ID_GROUP_ARRANGE_MONTH_HIGH:
'            frmInbox.CalendarControl.ViewType = xtpCalendarMonthView
'
'        Case ID_CALENDAREVENT_OPEN:
'            frmInbox.mnuOpenEvent
'
'        Case ID_CALENDAREVENT_DELETE:
'            frmInbox.mnuDeleteEvent
'
'        Case ID_CALENDAREVENT_NEW, ID_GROUP_NEW_APPOINTMENT:
'            'falta### frmEditEvent.AllDayOverride = False
'            frmInbox.mnuNewEvent
'            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
'
'        Case ID_GROUP_NEW_MEETING:
'            'falta### frmEditEvent.AllDayOverride = False
'            'falta### frmEditEvent.chkMeeting.Value = 1
'            frmInbox.mnuNewEvent
'            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
'
'        Case ID_GROUP_NEW_ALLDAY:
'            'falta### frmEditEvent.AllDayOverride = True
'            frmInbox.mnuNewEvent
'            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
'
'        Case ID_CALENDAREVENT_CHANGE_TIMEZONE:
'            frmInbox.mnuChangeTimeZone
'
'        Case ID_CALENDAREVENT_60:
'            frmInbox.mnuTimeScale 60
'
'        Case ID_CALENDAREVENT_30:
'            frmInbox.mnuTimeScale 30
'
'        Case ID_CALENDAREVENT_15:
'            frmInbox.mnuTimeScale 15
'
'        Case ID_CALENDAREVENT_10:
'            frmInbox.mnuTimeScale 10
'
'        Case ID_CALENDAREVENT_5:
'            frmInbox.mnuTimeScale 5
'
'
'
'
'        Case Else
'            AbiertoFormulario = True
'            AbrirFormularios Control.Id
'
'
'    End Select
'
'
'    If AbiertoFormulario Then
'        AbiertoFormulario = False
'        'mOTIVO... no lo se
'        'Pero si lo vamos cambiando funciona
'        If Me.DockingPaneManager.Panes(1).Enabled = 3 Then
'            Me.DockingPaneManager.Panes(1).Enabled = 3
'            Me.DockingPaneManager.Panes(2).Enabled = 3
'
'            frmPaneCalendar.DatePicker.Enabled = True
'
'            DockingPaneManager.RedrawPanes
'
'
'        Else
'            Me.DockingPaneManager.Panes(1).Enabled = 3
'            Me.DockingPaneManager.Panes(2).Enabled = 3
'
'        End If
'        DockingPaneManager.NormalizeSplitters
'
'    End If
'End Sub
'
'
'
'Private Sub CommandBars_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
'        Dim Control As CommandBarControl, ControlItem As CommandBarControl
'
'        If TypeOf CommandBar Is RibbonBackstageView Then
'            Debug.Print "RibbonBackstageView"
'        End If
'
'        Set Control = CommandBar.FindControl(, IDS_ARRANGE_BY)
'        If Not Control Is Nothing Then
'            Dim Index As Long
'            Index = Control.Index
'            Control.visible = False
'
'            Do While Index + 1 <= CommandBar.Controls.Count
'                Set ControlItem = CommandBar.Controls.Item(Index + 1)
'                If ControlItem.Id = IDS_ARRANGE_BY Then
'                    ControlItem.Delete
'                Else
'                    Exit Do
'                End If
'            Loop
'
''            Dim CurrentColumn As ReportColumn
''            For Each CurrentColumn In frmInbox. wndReportControl.Columns
''                Set ControlItem = CommandBar.Controls.Add(xtpControlButton, ID_REPORTCONTROL_COLUMN_ARRANGE_BY, CurrentColumn.Caption)
''                ControlItem.Parameter = CurrentColumn.ItemIndex
''                If Not frmInbox. wndReportControl.SortOrder.IndexOf(CurrentColumn) = -1 Then
''                    ControlItem.Checked = True
''                End If
''                If Not CurrentColumn.Visible Then
''                    ControlItem.Visible = False
''                End If
''            Next
'
'        End If
'End Sub
'
'Private Sub CommandBars_SpecialColorChanged()
'    Me.BackColor = CommandBars.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
'End Sub
'
'Private Sub CommandBars_ToolBarVisibleChanged(ByVal ToolBar As XtremeCommandBars.ICommandBar)
'     Debug.Print ToolBar.BarID
'End Sub

'Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
'
'    On Error Resume Next
'
'
'
'    Select Case Control.Id
'        Case ID_VIEW_STATUSBAR:
'            'Control.Checked = CommandBars.StatusBar.Visible
'
'
'
'        Case ID_GROUP_ARRANGE_WORK_WEEK:
'            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView, True, False)
'
'        Case ID_GROUP_ARRANGE_WEEK:
'            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWeekView, True, False)
'
'        Case ID_GROUP_ARRANGE_MONTH:
'            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarMonthView, True, False)
'
'        Case ID_OPTIONS_ANIMATION:
'            'Control.Checked = CommandBars.ActiveMenuBar.EnableAnimation
'
'        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
'             '   Dim newFontHeight As Integer
'             '   newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
'             '   Control.Checked = IIf(RibbonBar.FontHeight = newFontHeight, True, False)
'
'        Case ID_OPTIONS_FONT_AUTORESIZEICONS
'              '  Control.Checked = CommandBars.PaintManager.AutoResizeIcons
'
'        Case ID_RIBBON_EXPAND:
'            'Control.Visible = RibbonBar.Minimized
'
'        Case ID_RIBBON_MINIMIZE:
'            'Control.Visible = Not RibbonBar.Minimized
'    End Select
'    If Err.Number <> 0 Then Err.Clear
'End Sub

'Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
'    If (Action = PaneActionSplitterResized) Then
'        DockingPaneManager.RecalcLayout
'
'        ' Save MRUShortcutBarWidth
'        If (frmShortBar.ScaleWidth > MinimizedShortcutBarWidth And Container.Container.Type = PaneTypeSplitterContainer) Then
'            Debug.Print frmShortBar.ScaleWidth
'            MRUShortcutBarWidth = frmShortBar.ScaleWidth
'        End If
'    Else
'        If (Action = PaneActionSplitterResized) Then Debug.Print "Resizing "
'    End If
'End Sub
'
'
'Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
'    If Item.Tag = PANE_SHORTCUTBAR Then
'        Item.Handle = frmShortBar.hwnd
'    ElseIf Item.Tag = PANE_REPORT_CONTROL Then
'        Item.Handle = frmInbox.hwnd
'    End If
'End Sub
'
'
'
'Private Sub CreateCalendarTabOriginal()
'
'    Dim TabCalendarHome As RibbonTab
'    Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup
'
'
'    Dim Control As CommandBarControl
'    Dim ControlNew_NewItems As CommandBarPopup
'    Dim ControlArrange_Month As CommandBarPopup
'    Dim ControlManage_Open As CommandBarPopup
'    Dim ControlManage_Groups As CommandBarPopup
'    Dim ControlShare_Publish As CommandBarPopup
'
'    Dim PopupBar As CommandBar
'
'    Set TabCalendarHome = RibbonBar.InsertTab(14, "Agenda")
'    TabCalendarHome.Id = ID_TAB_CALENDAR_HOME
'
'    Set GroupNew = TabCalendarHome.Groups.AddGroup("&Nueva", ID_GROUP_NEW)
'
'    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Evento")
'    Control.Enabled = False
'    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_MEETING, "&Cita")
'    Control.Enabled = False
'
'    '------------------------------------
'    'Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
'    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "Evento")
'    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "E&vento todo el dia")
'    '    Control.BeginGroup = True
'    'ControlNew_NewItems.KeyboardTip = "V"
'
'    Set GroupGoTo = TabCalendarHome.Groups.AddGroup("I&r a", ID_GROUP_GOTO)
'    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_TODAY, "&Hoy")
'    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_NEXT7DAYS, "Próximos &7 dias ")
'    GroupGoTo.ShowOptionButton = True
'    GroupGoTo.ControlGroupOption.Caption = "Ir a (Ctrl+G)"
'    GroupGoTo.ControlGroupOption.ToolTipText = "Ir a (Ctrl+G)"
'    GroupGoTo.ControlGroupOption.DescriptionText = "Ir a fecha especificada."
'
'    Set GroupArrange = TabCalendarHome.Groups.AddGroup("Vista", ID_GROUP_ARRANGE2)
'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_DAY, "&Dia vista")
'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WORK_WEEK, "Samana &trabajo")
'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WEEK, "Sema&na vista")
'    Set ControlArrange_Month = GroupArrange.Add(xtpControlSplitButtonPopup, ID_GROUP_ARRANGE_MONTH, "Mes")
'            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_LOW, "Ver detalle")
'            Control.ToolTipText = "Muestra solo eventos todo el dia."
'            Control.DescriptionText = Control.ToolTipText
'            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_MEDIUM, "Detalle &Medio")
'            Control.ToolTipText = "Eventos todo el dia y si esta libre el dia o tiene eventos."
'            Control.DescriptionText = Control.ToolTipText
'            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_HIGH, "Detalle &Alto")
'            Control.ToolTipText = "Muestra todo."
'            Control.DescriptionText = Control.ToolTipText
'
''    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_SCHEDULE_VIEW, "Schedule View")
''    GroupArrange.ShowOptionButton = True
''    GroupArrange.ControlGroupOption.Caption = "Calendar Options"
''    GroupArrange.ControlGroupOption.ToolTipText = "Calendar Options"
''    GroupArrange.ControlGroupOption.DescriptionText = "Change the settings for calendars, meetings and time zones."
''
''
'
'
'End Sub
'
'
'
'Private Sub GuardarDatosUltimaTab()
'    i = RibbonBar.SelectedTab.Id
'    If i = ID_TAB_CALENDAR_HOME Then Exit Sub 'no guardo este tab
'    If i <> vUsu.TabPorDefecto Then
'        vUsu.TabPorDefecto = i
'        vUsu.GuardarTabPorDefecto
'    End If
'End Sub
'
'
'Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
'    Dim Control As CommandBarControl
'    Set Control = Controls.Add(ControlType, Id, Caption)
'
'    Control.BeginGroup = BeginGroup
'    Control.DescriptionText = DescriptionText
'    Control.Style = ButtonStyle
'    Control.Category = Category
'
'    Set AddButton = Control
'
'End Function
'
'Private Sub CommandBars_Resize()
'
'    On Error Resume Next
'
'    Dim Left As Long
'    Dim top As Long
'    Dim Right As Long
'    Dim Bottom As Long
'
'    CommandBars.GetClientRect Left, top, Right, Bottom
'
'End Sub
'
'Public Sub OpcionesMenuInformacion(Id As Long)
'
'    Select Case Id
'    Case ID_Licencia_Usuario_Final_txt
'        LanzaVisorMimeDocumento Me.hwnd, "c:\programas\Ariadna.rtf"
'    Case ID_Licencia_Usuario_Final_web
'        LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriTAXI-6.html?Licenciadeuso.html"
'    Case ID_Ver_Version_operativa_web
'        LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "Aritaxi-6.html"  ' "http://www.ariadnasw.com/clientes/"
'    End Select
'
'End Sub
'
'
'Private Sub LoadIcons()
'    CommandBars.Icons.RemoveAll
'    SuiteControlsGlobalSettings.Icons.RemoveAll
'    ReportControlGlobalSettings.Icons.RemoveAll
'
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\help.png", ID_APP_ABOUT, xtpImageNormal
'
'
'
'
'    'Para que no carge imagen de ratios y graficas y punteo, no lo pongo aqui ya que los cargo "pequeños"
'    '
'
'
'    'ICONOS PEQUEÑOS
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
'            Array(ID_RatiosyGráficas, ID_EvolucióndeSaldos, ID_Totalesporconcepto, 1, 1, ID_AseguClientes), xtpImageNormal
'
'
'
'
'    'Pequeños
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
'            Array(ID_ConsoBalSums, 1, 1, 1, ID_EstadísticaInmovilizado, ID_SimulaciónAmortización, ID_DeshacerAmortización, 1, 1, ID_VentaBajainmovilizado), xtpImageNormal
'
'    'Pequeños diario
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
'            Array(ID_TiposdeDiario, ID_TiposdePago, ID_ModelosdeCartas, ID_BicSwift, 1, ID_Agentes), xtpImageNormal
'
'         '      ID_AsientosPredefinidos
'
'    'Deberiamos cargar un array con unos(1) de longitud 143
'    ' y en funcion del valor del campo imagen en el punto de menu correspondiente
'    ' lo pondremos en el array.
'    ' Ejemplo    303 Extractos  Campo imagen: 87
'    ' quiere decir que en el campo 87 del array sustituieremos el 1 por el 303
'
'
''
'    Dim T() As Variant
'    'Cad linea son 15
'    T = Array(1, ID_Conceptos, ID_TiposdeIVA, ID_Bancos, ID_FormasdePago, ID_FacturasRecibidas, 1, ID_FacturasEmitidas, ID_LibroFacturasRecibidas, 1, 1, 1, 1, 1, 1, _
'        ID_RelaciónClientesporcuenta, ID_RelacionProveedoresporcuenta, 1, 1, 1, ID_RealizarCobro, ID_RealizarPago, 1, ID_Elementos, 1, 1, 1, 1, 1, ID_Punteoextractobancario, _
'        1, ID_InformePagospendientes, 1, 1, ID_Empresa, ID_ParametrosContabilidad, 1, ID_Contadores, ID_Extractos, ID_CarteradePagos, 1, 1, 1, ID_Punteo, 1, _
'        1, ID_PlanContable, 1, 1, 1, ID_Informes, 1, ID_Usuarios, 1, 1, 1, 1, ID_Nuevaempresa, ID_ConfigurarBalances, 1, _
'        1, ID_Compensaciones, 1, 1, 1, 1, ID_ConceptosInm, 1, 1, ID_GenerarAmortización, ID_Reclamaciones, 1, 1, 1, 1, _
'        ID_ImportarFacturasCliente, 1, 1, ID_Compensarcliente, ID_SumasySaldos, ID_CuentadeExplotación, ID_BalancedeSituación, ID_PérdidasyGanancias, 1, 1, 1, 1, 1, ID_CarteradeCobros, ID_InformeCobrosPendientes, _
'        ID_Renumeracióndeasientos, ID_CierredeEjercicio, ID_Deshacercierre, ID_DiarioOficial, ID_PresentaciónTelemáticadeLibros, ID_Traspasodecuentasenapuntes, ID_Renumerarregistrosproveedor, ID_TraspasocodigosdeIVA, 1, 1, 1, 1, 1, 1, 1, _
'        ID_Traspasodecuentasenapuntes, ID_Aumentardígitoscontables, 1, 1, 1, 1, 1, ID_LibroFacturasEmitidas, 1, 1, ID_Remesas, 1, 1, 1, 1, _
'        ID_RecepcionTalónPagare, ID_RemesasTalenPagare, ID_Accionesrealizadas, 1, ID_LiquidacionIVA, 1, 1, 1, ID_AsientosPredefinidos, 1, 1, 1, 1, ID_FrasConso, 1, _
'        ID_Renumerarregistrosproveedor, ID_TraspasocodigosdeIVA, 1, 1, 1, 1, ID_Asientos, 1)
'
'
'
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", T, xtpImageNormal
'
'
'
'    'Este de abjo funciona correctamente.
'    'NO tocar. Es por si falla volver a empezar
''    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", _
''            Array(ID_CarteradeCobros, ID_InformeCobrosPendientes, ID_RealizarCobro, ID_Compensarcliente, 1, ID_BalancePresupuestario, 1, _
''            ID_CentrosdeCoste, 1, 1, ID_Presupuestos, ID_Remesas, ID_Detalledeexplotación, ID_CarteradePagos, ID_CuentadeExplotaciónAnalítica, ID_ExtractosporCentrodeCoste, _
''            ID_Asientos, ID_Extractos, ID_Punteo, 1, ID_CuentadeExplotación, ID_Totalesporconcepto, ID_BalancedeSituación, ID_PérdidasyGanancias, _
''            ID_SumasySaldos, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
''            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
''            ID_Empresa, ID_ParametrosContabilidad, ID_Contadores, ID_Usuarios, 1, ID_Informes, ID_Nuevaempresa, ID_ConfigurarBalances, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
''            ID_FacturasEmitidas, ID_LibroFacturasEmitidas, ID_FacturasRecibidas, ID_LibroFacturasRecibidas, 1, 1, 1, 1, 1, ID_Elementos, ID_GenerarAmortización, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
''            1, ID_PlanContable, ID_TiposdeDiario, ID_Conceptos, ID_TiposdeIVA, ID_TiposdePago, ID_Bancos, ID_FormasdePago, _
''            ID_BicSwift, ID_Agentes, ID_AsientosPredefinidos, ID_ModelosdeCartas, _
''            ID_Renumeracióndeasientos, ID_CierredeEjercicio, ID_Deshacercierre, 1, 1, 1, 1, 1, 1, ID_DiarioOficial, _
''            ID_PresentaciónTelemáticadeLibros, ID_Traspasodecuentasenapuntes, ID_Renumerarregistrosproveedor, 1, ID_TraspasocodigosdeIVA), xtpImageNormal
''
'
'    'Presupuiestaria y analitaica cargadas arriba en pequeño
'    '---------------------------------------------------------
'    '
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
'            Array(ID_CentrosdeCoste, ID_ExtractosporCentrodeCoste, ID_Detalledeexplotación, ID_CuentadeExplotaciónAnalítica, ID_Presupuestos, ID_BalancePresupuestario), xtpImageNormal
'
'
'
'
'    'Pequeños
'    ' ID_Compensaciones ID_Reclamaciones  ID_InformeImpagados ID_RemesasTalenPagare ID_Norma57Pagosventanilla  ID_TransferenciasAbonos
'    ' ID_InformePagosbancos ID_Transferencias ID_Pagosdomiciliados ID_GastosFijos ID_Compensarproveedor ID_Confirming
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
'            Array(1, ID_Reclamaciones, ID_InformeImpagados, ID_RemesasTalenPagare, ID_Norma57Pagosventanilla, ID_TransferenciasAbonos, ID_Confirming, _
'            ID_Pagosdomiciliados, ID_GastosFijos, ID_Compensarproveedor), xtpImageNormal
'
'
'
'
'
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
'            Array(ID_InformePagosbancos, ID_Transferencias, ID_MemoriaPlazosdepago, ID_Informeporcuenta, ID_SituaciónTesoreria, ID_InformeporNIF), xtpImageNormal
'
'
'
'
'
'    '------------------------------------------------------------------------------------------------------------------------
'    '------------------------------------------------------------------------------------------------------------------------
'    '------------------------------------------------------------------------------------------------------------------------
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookcalicons.png", _
'            Array(ID_GROUP_NEW_APPOINTMENT, ID_GROUP_NEW_MEETING, ID_GROUP_NEW_ITEMS, ID_GROUP_GOTO_TODAY, _
'            ID_GROUP_GOTO_NEXT7DAYS, ID_GROUP_ARRANGE_DAY, ID_GROUP_ARRANGE_WORK_WEEK, ID_GROUP_ARRANGE_WEEK, _
'            ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_SCHEDULE_VIEW, ID_GROUP_MANAGE_CALENDARS_OPEN, ID_GROUP_MANAGE_CALENDARS_GROUPS, _
'            ID_GROUP_SHARE_EMAIL, ID_GROUP_SHARE_SHARE, ID_GROUP_SHARE_PUBLISH, ID_GROUP_SHARE_PERMISSIONS), xtpImageNormal
'
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\RibbonMinimize.png", _
'            Array(ID_RIBBON_MINIMIZE, ID_RIBBON_EXPAND), xtpImageNormal
'
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\Search.png", _
'            ID_SEARCH_ICON, xtpImageNormal
'
'     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonslarge.png", _
'            Array(ID_GROUP_MAIL_NEW_NEW, ID_GROUP_MAIL_NEW_NEW_ITEMS, ID_GROUP_MAIL_DELETE_DELETE, ID_GROUP_MAIL_RESPOND_REPLY, _
'            ID_GROUP_MAIL_RESPOND_REPLY_ALL, ID_GROUP_MAIL_RESPOND_FORWARD, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
'
'     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", _
'            Array(ID_GROUP_MAIL_DELETE_CLEANUP, ID_GROUP_MAIL_DELETE_JUNK, ID_GROUP_MAIL_RESPOND_MEETING, ID_GROUP_MAIL_RESPOND_IM, _
'            ID_GROUP_MAIL_RESPOND_MORE, ID_GROUP_MAIL_TAGS_UNREAD, ID_GROUP_MAIL_TAGS_CATEGORIZE, ID_GROUP_MAIL_TAGS_FOLLOWUP, ID_GROUP_MAIL_FIND_ADDRESSBOOK, _
'            ID_GROUP_MAIL_FIND_FILTER, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
'
'        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookpane.png", _
'            Array(ID_SWITCH_NORMAL, ID_SWITCH_CALENAR_AND_TASK, ID_SWITCH_CALENDAR, ID_SWITCH_CLASSIC, ID_SWITCH_READING), xtpImageNormal
'
'        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
'            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
'            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
'        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_24x24.bmp", _
'            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
'            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
'
'        CommandBars.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
'            Array(ID_QUICKSTEP_REPLAY_DELETE, ID_QUICKSTEP_TO_MANAGER, ID_QUICKSTEP_MOVE_TO, ID_QUICKSTEP_CREATE_NEW, ID_QUICKSTEP_TEAM_EMAIL, ID_QUICKSTEP_DONE), xtpImageNormal
'
'        ReportControlGlobalSettings.Icons.LoadBitmap App.Path & "\styles\bmreport.bmp", _
'        Array(COLUMN_MAIL_ICON, COLUMN_IMPORTANCE_ICON, COLUMN_CHECK_ICON, RECORD_UNREAD_MAIL_ICON, RECORD_READ_MAIL_ICON, _
'            RECORD_REPLIED_ICON, RECORD_IMPORTANCE_HIGH_ICON, COLUMN_ATTACHMENT_ICON, COLUMN_ATTACHMENT_NORMAL_ICON, _
'            RECORD_IMPORTANCE_LOW_ICON), xtpImageNormal
'
'
'        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\suministro-inmediato-informacion.bmp", ID_SII, xtpImageNormal
'
'
'        Dim i As Integer
'        For i = 1 To 17
'            SuiteControlsGlobalSettings.Icons.LoadIcon App.Path & "\styles\TreeView\icon" & i & ".ico", i, xtpImageNormal
'        Next i
'End Sub
'
'
''Establecer y fijar Skin
'Public Sub EstablecerSkin(QueSkin As Integer)
'
'    FijaSkin QueSkin
'
'  ' Cargando el archivo del Skin
'  ' ============================
'    'frmPpal2.SkinFramework1.LoadSkin Skn$, ""
'    Me.SkinFramework1.ApplyWindow frmPpal2.hwnd
'    Me.SkinFramework1.ApplyOptions = Me.SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
'
'
'
'
'End Sub
'
'Private Function FijaSkin(numero)
'    Me.SkinFramework1.ExcludeModule "crviewer9.dll"
'
'  Select Case (numero)
'
'
'            Case 1:
'                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
'                Me.SkinFramework1.LoadSkin Skn$, "NormalBlue.ini"
'            Case 2:
'                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
'                Me.SkinFramework1.LoadSkin Skn$, "NormalSilver.ini"
'            Case 3:
'                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
'                Me.SkinFramework1.LoadSkin Skn$, "NormalBlack.ini"
'
'
'
'  End Select
'
'End Function

'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
'**************************************************************************************************************
Private Sub AbrirFormularios(Accion As Long)
    
   
   ' If Accion <> ID_SII Then AbrirFormSII_2 False
    
    
    Select Case Accion
        Case 101 ' empresa
'            frmempresa.Show vbModal
'        Case 102 ' parametros contabilidad
'            If Not (vEmpresa Is Nothing) Then
'                frmparametros.Show vbModal
'            End If
'        Case 103 ' parametros tesoreria
'        Case 104 ' contadores
'            Screen.MousePointer = vbHourglass
'            If vUsu.Nivel = 0 Then frmContadores.Show vbModal
'            Screen.MousePointer = vbDefault
'        Case 105 ' usuarios
'            frmMantenusu.Show vbModal
'        Case 106 ' informes
'            frmCrystal.Show vbModal
'        Case 107 ' crear nueva empresa
'            If vUsu.Nivel > 1 Then Exit Sub
'
'            frmCentroControl.Opcion = 2
'            frmCentroControl.Show vbModal
'        Case 108 'Configurar Balances
'            Screen.MousePointer = vbHourglass
'            frmColBalan.Show vbModal, Me
'
'        Case 199
'             frmCalendarCategorias.Show vbModal
'
'
'        Case 201 ' plan contable
'            Screen.MousePointer = vbHourglass
'            frmColCtas.ConfigurarBalances = 0
'            frmColCtas.DatosADevolverBusqueda = ""
'            frmColCtas.Show vbModal, Me
'        Case 202 ' tipos de diario
'            Screen.MousePointer = vbHourglass
'            frmTiposDiario.Show vbModal
'        Case 203 ' conceptos
'            Screen.MousePointer = vbHourglass
'            frmConceptos.Show vbModal
'        Case 204 ' tipos de iva
'            Screen.MousePointer = vbHourglass
'            frmIVA.Show vbModal
'        Case 205 ' tipos de pago
'            Screen.MousePointer = vbHourglass
'            frmTipoPago.Show vbModal
'        Case 206 ' formas de pago
'            Screen.MousePointer = vbHourglass
'            frmFormaPago.Show vbModal
'        Case 207 ' bancos
'            Screen.MousePointer = vbHourglass
'            frmBanco.Show vbModal
'        Case 208 ' bic
'            Screen.MousePointer = vbHourglass
'            frmBic.Show vbModal
'        Case 209 ' agentes
'            Screen.MousePointer = vbHourglass
'            frmAgentes.Show vbModal
'        Case 210 ' departamentos
'        Case 211 ' asientos predefinidos
'            Screen.MousePointer = vbHourglass
'            frmAsiPre.Show vbModal
'        Case 212 ' cartas de reclamacion
'            Screen.MousePointer = vbHourglass
'            frmCartas.Show vbModal
'
'        Case 301 ' asientos
'            Screen.MousePointer = vbHourglass
'            frmAsientosHco.Asiento = ""
'            frmAsientosHco.DesdeNorma43 = 0
'            frmAsientosHco.Show vbModal
'        Case 303 ' extractos
'            Screen.MousePointer = vbHourglass
'            frmConExtr.EjerciciosCerrados = False
'            frmConExtr.cuenta = ""
'            frmConExtr.Show vbModal
'        Case 304 ' punteo
'            Screen.MousePointer = vbHourglass
'            frmPuntear.EjerciciosCerrados = False
'            frmPuntear.Show vbModal
'        Case 305 ' reemision de diarios
''            AbrirListado 6, False
'        Case 306 ' sumas y saldos
'            frmInfBalSumSal.Show vbModal
'
'        Case 307 ' cuenta de explotacion
'            frmInfCtaExplo.Show vbModal
'
'        Case 308 ' balance de situacion
'            frmInfBalances.Opcion = 0
'            frmInfBalances.Show vbModal
'
'        Case 309 ' perdidas y ganancias
'            frmInfBalances.Opcion = 1
'            frmInfBalances.Show vbModal
'
'        Case 310 ' totales por concepto
'            frmInfTotCtaCon.Show vbModal
'        Case 311 ' evolucion de saldos
'            frmInfEvolSal.Show vbModal
'        Case 312 ' ratios y graficas
'            frmInfRatios.Show vbModal
'        Case 314 ' puntero extracto bancario
'            frmPunteoBanco.Show vbModal
'
'        Case 315
'            frmInfBalSumSalConso.Show vbModal
'
'        Case 401 ' emitidas
'            Screen.MousePointer = vbHourglass
'            frmFacturasCli.FACTURA = ""
'            frmFacturasCli.Show vbModal
'        Case 402 ' libro emitidas
'            frmFacturasCliListado.Show vbModal
'        Case 403 ' relacion clientes por cuenta
'            frmFacturasCliCtaVtas.Show vbModal
'        Case 404 ' recibidas
'            Screen.MousePointer = vbHourglass
'            frmFacturasPro.Show vbModal
'        Case 405 ' libro recibidas
'            frmFacturasProListado.Show vbModal
'        Case 406 ' relacion proveedores por cuenta
'            frmFacturasProCtaGastos.Show vbModal
'        Case 407 ' liquidacion iva
''            AbrirListado 12, False
'        Case 408 ' certificado iva
'            frmModelo303.OpcionListado = 0
'            frmModelo303.Show vbModal
'        Case 409 ' modelo 340
'            frmModelo340.Show vbModal
'        Case 410 ' modelo 347
'            frmModelo347.Show vbModal
'        Case 411 ' modelo 349
'            frmModelo349.Show vbModal
'        Case 412 ' liquidacion de iva
'            frmHcoLiqIVA.Show vbModal
'
'        Case ID_FrasConso   '413
'            frmConsolidadoFras.Show vbModal
'
'        Case ID_AseguClientes   '414
'            frmSegurosListClientes.Show vbModal
'
'        Case ID_AseguComunicaSeguro '415
'            frmSegurosListComunicacion.Show vbModal
'
'        Case ID_SII
'            AbrirFormSII_2 True
'
'
'
'        Case 502 ' conceptos
'            Screen.MousePointer = vbHourglass
'            frmInmoConceptos.Show vbModal
'        Case 503 ' elementos
'            frmInmoElto.DatosADevolverBusqueda = ""
'            frmInmoElto.Show vbModal
'        Case 505 ' estadistica
'            frmInmoInfEst.Show vbModal
'        Case 507 ' historico inmovilizado
'            Screen.MousePointer = vbHourglass
'            frmInmoHco.Show vbModal
'        Case 508 ' simulacion
'            frmInmoSimu.Show vbModal
'        Case 509 ' calculo y contabilizacion
'            frmInmoGenerar.Opcion = 2
'            frmInmoGenerar.Show vbModal
'        Case 510 ' deshacer amortizacion
'            frmInmoDeshacer.Show vbModal
'        Case 511 ' venta-baja inmmovilizado
'            frmInmoVenta.Opcion = 3
'            frmInmoVenta.Show vbModal
'        Case 601 ' cartera de cobros
'            frmTESCobros.Show vbModal
'        Case 602 ' informe de cobros pendientes
'            frmTESCobrosPdtesList.Show vbModal
'
'        Case 603 ' impresion de recibos
'            frmTESImpRecibo.documentoDePago = ""
'            frmTESImpRecibo.Show vbModal
'        Case 604 ' realizar cobro
'            With frmTESRealizarCobros
'
'                '--.vSQL = SQL
'                .Regresar = False
'                .Cobros = True
'                .ContabTransfer = False
'                .SegundoParametro = ""
'                'Los textos
''                .vTextos = Text1(2).Text & "|" & Me.txtCta(0).Text & " - " & Me.txtDescCta(0).Text & "|" & SubTipo & "|"
'
'                'Marzo2013   Cobramos un solo cliente
'                'Aparecera un boton para traer todos los cobros
'                '.CodmactaUnica = "4300000001" 'Trim(txtCtaNormal(9).Text)
'                .Show vbModal
'            End With
'
'        Case 606 ' compensaciones
'            frmTESCompensaciones.Show vbModal
'        Case 607 ' compensar cliente
'            CadenaDesdeOtroForm = ""
'            frmTESCompensaAboCli.Show vbModal
'        Case 608 ' reclamaciones
'            frmTESReclamaCli.Show vbModal
'        Case 609 ' remesas
'            frmTESRemesas.Tipo = 1 ' efectos
'            frmTESRemesas.Show vbModal
'        Case 610 ' Informe Impagados
'            frmTESCobrosDevList.Show vbModal
'        Case 611 ' Recepción Talón-Pagaré
'            frmTESRecepcionDoc.Show vbModal
'        Case 612 ' Remesas Talón-Pagaré
'            frmTESRemesasTP.Tipo = 2 ' talon pagare
'            frmTESRemesasTP.Show vbModal
'
'        Case 613 ' Norma 57: Pago por ventanilla
'            frmTESNorma57.Opcion = 42
'            frmTESNorma57.Show vbModal
'
'        Case 614 ' transferencia abonos
'            frmTESTransferencias.TipoTrans = 0 ' de abonos
'            frmTESTransferencias.Show vbModal
'
'        Case 709 ' Abono remesa
'        Case 710 ' Devoluciones
'        Case 711 ' Eliminar riesgo
'
'        Case 801 ' Cartera de Pagos
'            frmTESPagos.Show vbModal
'        Case 802 ' Informe Pagos pendientes
'            frmTESPagosPdtesList.Show vbModal
'        Case 803 ' Informe Pagos bancos
'            frmTESPagosBancoList.Show vbModal
'        Case 804 ' Realizar Pago
'            frmTESRealizarPagos.Show vbModal
'        Case 805 ' Transferencias
'            frmTESTransferencias.TipoTrans = 1 ' de pagos
'            frmTESTransferencias.Show vbModal
'        Case 806 ' Pagos domiciliados
'            frmTESTransferencias.TipoTrans = 2 ' pagos domiciliados
'            frmTESTransferencias.Show vbModal
'
'        Case 807 ' Gastos Fijos
'            frmTESGastosFijos.Show vbModal
'
'        Case 808 ' Memoria Pagos proveedores
'
'        Case 809 ' Compensar proveedor
'            CadenaDesdeOtroForm = ""
'            frmTESCompensaAboPro.Show vbModal
'
'        Case 810 ' Confirming
'            frmTESTransferencias.TipoTrans = 3 ' confirming
'            frmTESTransferencias.Show vbModal
'
'        Case 901 ' Informe por NIF
'            frmTESInfSituacionNIF.Show vbModal
'
'        Case 902 ' Informe por cuenta
'            frmTESInfSituacionCta.Show vbModal
'
'        Case 903 ' Situación Tesoreria
'            frmTESInfSituacion.Show vbModal
'
'
'        ' Analitica
'        Case 1001 ' Centros de Coste
'            frmCCCentroCoste.Show vbModal
'
'        Case 1002 ' Consulta de Saldos
'            frmCCConExtr.Show vbModal
'
'        Case 1003 ' Cuenta de Explotación
'            frmCCCtaExplo.Show vbModal
'        Case 1004 ' Centros de coste por cuenta
'            AbrirListado 17, False
'        Case 1005 ' Detalle de explotación
'            frmCCDetalleExplota.Show vbModal
'
'        ' Presupuestaria
'        Case 1101 ' Presupuestos
'            Screen.MousePointer = vbHourglass
'            'frmColPresu.Show vbModal
'            frmPresu.Show vbModal
'        Case 1102 ' Listado de Presupuestos
''            AbrirListado 9, False
'        Case 1103 ' Balance Presupuestario
'            frmPresuBal.Show vbModal
'
'        ' Consolidado
'        Case 1201 ' Sumas y Saldos
'            AbrirListado 24, False
'        Case 1202 ' Balance de Situación
'            AbrirListado 51, False
'        Case 1203 ' Pérdidas y Ganancias
'            AbrirListado 50, False
'        Case 1204 ' Cuenta de Explotación
'            AbrirListado 31, False
'        Case 1205 ' Listado Facturas Clientes
'            AbrirListado 53, False
'        Case 1206 ' Listado Facturas Proveedores
'            AbrirListado 52, False
'
'        ' Cierre de Ejercicio
'        Case 1301 ' Renumeración de asientos
'            frmCierre.Opcion = 0
'            frmCierre.Show vbModal
'        Case 1302 ' Simulación de cierre
'            frmCierre.Opcion = 4
'            frmCierre.Show vbModal
'        Case 1303 ' Cierre de Ejercicio
'            frmCierre.Opcion = 1
'            frmCierre.Show vbModal
'        Case 1304 ' Deshacer cierre
'            frmCierre.Opcion = 5
'            frmCierre.Show vbModal
'        Case 1305 ' Diario Oficial
''            AbrirListado 14, False
'        Case 1306 ' Diario Oficial Resumen
''            AbrirListado 18, False
'            frmInfDiarioOficial.Show vbModal
'        Case 1307 ' Presentación cuentas anuales
'            Telematica 0
'        Case 1308 ' Presentación Telemática de Libros
'            Telematica 1
'        Case 1309 ' memoria de Plazos de Pago
'            frmTESMemoriaPlazos.Show vbModal
'
'        ' Utilidades
'        Case 1401 ' Comprobar cuadre
'            Screen.MousePointer = vbHourglass
'            frmMensajes.Opcion = 2
'            frmMensajes.Show vbModal
'        Case 1403 ' Revisar caracteres especiales
''            Screen.MousePointer = vbHourglass
''            frmMensajes.opcion = 14
''            frmMensajes.Show vbModal
'
'        Case 1404 ' Agrupacion cuentas
'        Case 1405 'Buscar ...
'
'        Case 1407 'Desbloquear asientos
'            mnHerrAriadnaCC_Click (0)
'        Case 1408 'Mover cuentas
'            mnHerrAriadnaCC_Click (1)
'        Case 1409 'Renumerar registros proveedor
'            mnHerrAriadnaCC_Click (5)
'        Case 1410 'Aumentar dígitos contables
'            mnHerrAriadnaCC_Click (3)
'        Case 1411 'cambio de iva
'            mnHerrAriadnaCC_Click (4)
'        Case 1412 'log de acciones
'            Screen.MousePointer = vbHourglass
'            Load frmLog
'            DoEvents
'            frmLog.Show vbModal
'            Screen.MousePointer = vbDefault
'        Case 1413
'            frmImportarFraCli.Show vbModal
'        Case 1414
'            frmImportarNavarres.Show vbModal
'
'        Case Else
'
    End Select
     
     
     
   If Timer - UltimaLecturaReminders > 300 Then
        frmReminders.OnReminders xtpCalendarRemindersFire, Nothing
        If frmReminders.CuantosAvisos > 0 Then frmReminders.Show vbModal, Me
        CerrarAvisos
        UltimaLecturaReminders = Timer
    End If
     
End Sub

Private Sub CerrarAvisos()
    On Error Resume Next
    Unload frmReminders
    Err.Clear
End Sub



