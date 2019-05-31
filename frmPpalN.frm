VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "CODEJO~3.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#17.2#0"; "COA2AE~1.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#17.2#0"; "COC9F8~1.OCX"
Begin VB.Form frmppal 
   Caption         =   "Aritaxi"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11580
   FillStyle       =   0  'Solid
   Icon            =   "frmPpalN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlToolbarIcons_N 
      Left            =   9090
      Top             =   3930
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7532
            Key             =   "New"
            Object.Tag             =   "100"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7590
            Key             =   "Open"
            Object.Tag             =   "101"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":75EE
            Key             =   "Save"
            Object.Tag             =   "103"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":764C
            Key             =   "Print"
            Object.Tag             =   "113"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":76AA
            Key             =   "Cut"
            Object.Tag             =   "108"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7708
            Key             =   "Copy"
            Object.Tag             =   "106"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7766
            Key             =   "Paste"
            Object.Tag             =   "107"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":77C4
            Key             =   "Bold"
            Object.Tag             =   "120"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7822
            Key             =   "Italic"
            Object.Tag             =   "121"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7880
            Key             =   "Underline"
            Object.Tag             =   "122"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":78DE
            Key             =   "Align Left"
            Object.Tag             =   "123"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":793C
            Key             =   "Center"
            Object.Tag             =   "124"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":799A
            Key             =   "Align Right"
            Object.Tag             =   "125"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":79F8
            Key             =   "About"
            Object.Tag             =   "112"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7A56
            Key             =   ""
            Object.Tag             =   "166"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7AB4
            Key             =   ""
            Object.Tag             =   "168"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7B12
            Key             =   ""
            Object.Tag             =   "165"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListPPal48 
      Left            =   7590
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8460
      Top             =   5310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM_N 
      Left            =   9090
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN_N 
      Left            =   9870
      Top             =   3990
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN16 
      Left            =   6060
      Top             =   6210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM16 
      Left            =   6810
      Top             =   6210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListPpal16 
      Left            =   7560
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgListComun1 
      Left            =   5160
      Top             =   3510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   360
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32_N 
      Left            =   9840
      Top             =   5970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":E3D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":14C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1B496
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":21CF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":2855A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":2EDBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras_N 
      Left            =   10500
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":3561E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":3BE80
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":426E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":48F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":4F7A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":56008
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":5C86A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":630CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras_BN_N 
      Left            =   9870
      Top             =   5370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483626
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":63ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":6A340
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":70BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":77404
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":7DC66
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":844C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":8AD2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":97DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9E650
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2_N 
      Left            =   10560
      Top             =   4620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":9F062
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":A58C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":A8076
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListDocumentos_N 
      Left            =   10470
      Top             =   6000
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
            Picture         =   "frmPpalN.frx":AE8D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":AFB5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B230C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B4446
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B4760
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B7B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":B9764
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BA541
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BB4B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BC428
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImaListBotoneras32_BN_N 
      Left            =   9120
      Top             =   5340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":BD3C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":C3C27
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":CA489
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":D0CEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":D754D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":DDDAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":E4611
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListviews_N 
      Left            =   8400
      Top             =   4620
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
            Picture         =   "frmPpalN.frx":EAE73
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F16D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F3E87
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":F9AA9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcoForms_N 
      Left            =   9150
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":10030B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":100D1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":100DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun16_N 
      Left            =   8490
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_BN 
      Left            =   5190
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun_OM 
      Left            =   5190
      Top             =   5550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListComun16 
      Left            =   5190
      Top             =   6210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgIcoForms 
      Left            =   6030
      Top             =   4950
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
            Picture         =   "frmPpalN.frx":1017CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1021DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":102277
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":102C89
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1094EB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListPpal 
      Left            =   6030
      Top             =   4260
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
            Picture         =   "frmPpalN.frx":10FD4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":110DDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":111E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":112F03
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":113F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":115A17
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":116AA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":117B3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":118BCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":119C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":11ACF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":11BD83
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":11CE15
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":11DEA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":11EF39
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":11FFCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":12105D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1220EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":123181
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":124B13
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":12B375
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":12F877
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":130289
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":13367B
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":139EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":14073F
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1417D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":142863
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1438F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":144987
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":145A19
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":14C27B
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":14D30D
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":153B6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":154C01
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":155C93
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":156D25
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":157DB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListTPV 
      Left            =   6810
      Top             =   4260
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
            Picture         =   "frmPpalN.frx":158E49
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":15A7DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":15C16D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":15DAFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":15F491
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":160E23
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":1627B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":164147
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":16A9A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":17019B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListB 
      Left            =   6090
      Top             =   3480
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
            Picture         =   "frmPpalN.frx":1769FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":17740F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":177E21
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":178833
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":179245
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":179C57
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":17A669
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":17B07B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":17BA8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpalN.frx":17C49F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   390
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   5640
      Top             =   1080
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   4800
      Top             =   1920
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":17CEB1
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   3840
      Top             =   600
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   4320
      Top             =   1320
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeCommandBars.ImageManager ImageManagerGalleryStyles 
      Left            =   3360
      Top             =   120
      _Version        =   1114114
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPpalN.frx":17CECB
   End
End
Attribute VB_Name = "frmppal"
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



Public Function RibbonBar() As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
End Function

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

Public Sub CheckButton(nButton As Integer)
    CommandBars.Actions(ID_OPTIONS_STYLEBLUE2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLESILVER2010).Checked = False
    CommandBars.Actions(ID_OPTIONS_STYLEBLACK2010).Checked = False
    
    CommandBars.Actions(nButton).Checked = True
End Sub

Sub OnThemeChanged(Id As Integer)
Dim N_Skin As Integer
    CheckButton Id
    
    Dim FlatStyle As Boolean
    FlatStyle = Id >= ID_OPTIONS_STYLESCENIC7 And Id <= ID_OPTIONS_STYLEBLACK2010
        
        
    Me.BackColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
   
    
    CommandBars.EnableOffice2007Frame False

    Select Case CommandBars.VisualTheme
        Case xtpThemeResource, xtpThemeRibbon
            CommandBars.AllowFrameTransparency False 'True
            CommandBars.EnableOffice2007Frame True
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
        Case Else
            CommandBars.AllowFrameTransparency True
            CommandBars.EnableOffice2007Frame False
            CommandBars.SetAllCaps False
            CommandBars.statusBar.SetAllCaps False
    End Select
    
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CommandBars.ToolTipContext
    ToolTipContext.Style = xtpToolTipResource
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    
    statusBar.ToolTipContext.Style = ToolTipContext.Style
    frmShortBar.wndShortcutBar.ToolTipContext.Style = ToolTipContext.Style
    
       
    'CreateBackstage
    'SetBackstageTheme
    
    'CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\styles\Ariconta.ico", 16, 16
            
    'Set Captions VisualTheme
    On Error Resume Next
    Dim CtrlCaption As ShortcutCaption
    Dim Form As Form, Ctrl As Object
            
    For Each Form In Forms
        For Each Ctrl In Form.Controls
                    
            Set CtrlCaption = Ctrl
            If Not CtrlCaption Is Nothing Then
                CtrlCaption.VisualTheme = frmShortBar.wndShortcutBar.VisualTheme
            End If
                    
        Next
    Next
       
    DockingPaneManager.PaintManager.SplitterSize = 5
    DockingPaneManager.PaintManager.SplitterColor = frmShortBar.wndShortcutBar.PaintManager.SplitterBackgroundColor
    
    DockingPaneManager.PaintManager.ShowCaption = False
    DockingPaneManager.RedrawPanes
        
    frmShortBar.SetColor Id
    frmInbox.SetColor Id
        

    frmPaneCalendar.SetFlatStyle FlatStyle
    frmPaneContacts.SetFlatStyle FlatStyle
    'frmPaneInformacion.SetFlatStyle FlatStyle
    'frmPaneAcercaDe.SetFlatStyle FlatStyle
    
    
    
    
    
    
    LoadIcons
    N_Skin = Id - 2895
    EstablecerSkin N_Skin
    
    'Updatear SKIN usuario
    If CStr(N_Skin) <> vUsu.Skin Then
        vUsu.Skin = N_Skin
        vUsu.ActualizarSkin
    End If
    
End Sub

Public Sub SetBackstageTheme()
Dim i As Integer
    Dim nTheme As XtremeCommandBars.XTPBackstageButtonControlAppearanceStyle
    nTheme = xtpAppearanceResource

   ' If Not (pageBackstageInfo Is Nothing) Then
        'pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        'pageBackstageInfo.btnProtectDocument.Appearance = nTheme
        'pageBackstageInfo.btnCheckForIssues.Appearance = nTheme
        'pageBackstageInfo.btnManageVersions.Appearance = nTheme
   ' End If
    
    If Not (pageBackstageHelp Is Nothing) Then
        For i = 0 To 4
            pageBackstageHelp.btnAcciones(i).Appearance = nTheme
        Next
        
    End If
    
    'If Not (pageBackstageSend Is Nothing) Then
        'pageBackstageSend.btnTab(0).Appearance = nTheme
        'pageBackstageSend.btnTab(1).Appearance = nTheme
        'pageBackstageSend.btnTab(2).Appearance = nTheme
        'pageBackstageSend.btnTab(3).Appearance = nTheme
    'End If

End Sub

Private Sub CreateStatusBar()
Dim Pane As StatusBarPane

    If RibbonSeHaCreado Then
        'StatusBar.Pane(0).Value = vEmpresa.nomempre & "    " & vUsu.Login
        statusBar.Pane(0).Text = "Nº " & vEmpresa.codempre
        statusBar.Pane(1).Text = vEmpresa.nomempre
    
    Else
    
         
         Set statusBar = Nothing
         
         Set statusBar = CommandBars.statusBar
         statusBar.visible = True
         
         
         Set Pane = statusBar.AddPane(ID_INDICATOR_PAGENUMBER)
         Pane.Text = "Nº " & vEmpresa.codempre
         Pane.Caption = "&C"
         Pane.Value = vEmpresa.nomempre & "    " & vUsu.Login
         Pane.Button = True
         Pane.SetPadding 8, 0, 8, 0
         
         Set Pane = statusBar.AddPane(ID_INDICATOR_WORDCOUNT)
         Pane.Text = vEmpresa.nomempre
         Pane.Caption = ""
         Pane.Value = vEmpresa.codempre
         Pane.Button = True
         Pane.SetPadding 8, 0, 8, 0
         
         
         Set Pane = statusBar.AddPane(0)
         Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
         Pane.BeginGroup = True
                 
        '
         statusBar.RibbonDividerIndex = 3
         statusBar.EnableCustomization True
         
         CommandBars.Options.KeyboardCuesShow = xtpKeyboardCuesShowNever
         CommandBars.Options.ShowKeyboardTips = True
         CommandBars.Options.ToolBarAccelTips = True
    End If
End Sub

Private Sub DockBarRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.RecalcLayout
    BarOnLeft.GetWindowRect Left, top, Right, Bottom
    
    CommandBars.DockToolBar BarToDock, Right, (Bottom + top) / 2, BarOnLeft.Position

End Sub

Private Sub CommandBars_CommandBarKeyDown(CommandBar As XtremeCommandBars.ICommandBar, KeyCode As Long, Shift As Integer)
    Debug.Print CommandBar.BarID
End Sub

Public Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim AbiertoFormulario  As Boolean
    AbiertoFormulario = False
    

    Select Case Control.Id
        Case XTPCommandBarsSpecialCommands.XTP_ID_RIBBONCONTROLTAB:
            
        
          
        Case XTP_ID_RIBBONCUSTOMIZE:
            CommandBars.ShowCustomizeDialog 3
            
        Case ID_APP_ABOUT:
          
           LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriTAXI-6.html?"
   
        
        Case ID_FILE_NEW:
            'frmEmail.Show 0, Me
        
        
        
        Case ID_Licencia_Usuario_Final_txt, ID_Licencia_Usuario_Final_web, ID_Ver_Version_operativa_web
            OpcionesMenuInformacion Control.Id
        
        
        
        Case ID_VIEW_STATUSBAR:
            CommandBars.statusBar.visible = Not CommandBars.statusBar.visible
            CommandBars.RecalcLayout
            
        Case ID_RIBBON_EXPAND:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            RibbonBar.Minimized = Not RibbonBar.Minimized
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
            Dim newFontHeight As Integer
            newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
            RibbonBar.FontHeight = newFontHeight
            
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
            CommandBars.PaintManager.AutoResizeIcons = Not CommandBars.PaintManager.AutoResizeIcons
            CommandBars.RecalcLayout
            RibbonBar.RedrawBar
            
        Case ID_OPTIONS_STYLEBLUE2010:
            LoadResources "Office2010.dll", "Office2010Blue.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLUE2010
            
            
            
       Case ID_OPTIONS_STYLESILVER2010:
            LoadResources "Office2010.dll", "Office2010Silver.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLESILVER2010
        
       Case ID_OPTIONS_STYLEBLACK2010:
            LoadResources "Office2010.dll", "Office2010Black.ini"
            CommandBars.VisualTheme = xtpThemeRibbon
            DockingPaneManager.VisualTheme = ThemeResource
            frmShortBar.wndShortcutBar.VisualTheme = xtpShortcutThemeResource
            frmInbox.CalendarControl.VisualTheme = xtpCalendarThemeResource
            frmInbox.ScrollBarCalendar.Appearance = xtpAppearanceResource
            
            OnThemeChanged ID_OPTIONS_STYLEBLACK2010
        
        Case ID_APP_EXIT:
            Unload Me
        
    
            
        Case ID_GROUP_GOTO_TODAY:
            Select Case frmInbox.CalendarControl.ViewType
                Case xtpCalendarDayView:
                    frmInbox.CalendarControl.DayView.ShowDay DateTime.Now, True
            
                Case xtpCalendarWorkWeekView:
                    frmInbox.CalendarControl.DayView.SetSelection DateTime.Now, DateTime.Now, True
                    frmInbox.CalendarControl.RedrawControl
            
                Case xtpCalendarWeekView:
                    frmInbox.CalendarControl.WeekView.SetSelection DateTime.Now, DateTime.Now, True
            
                Case xtpCalendarMonthView:
                    frmInbox.CalendarControl.MonthView.SetSelection DateTime.Now, DateTime.Now, True
            End Select
            
        Case ID_GROUP_GOTO_NEXT7DAYS:
            Dim lastDate As Date
            lastDate = frmInbox.CalendarControl.DayView.Days(frmInbox.CalendarControl.DayView.DaysCount - 1).Date
            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
            frmInbox.CalendarControl.DayView.ShowDays lastDate + 1, lastDate + 7
            
        Case ID_GROUP_ARRANGE_DAY:
            frmInbox.CalendarControl.ViewType = xtpCalendarDayView
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView
            
        Case ID_GROUP_ARRANGE_WEEK:
            frmInbox.CalendarControl.UseMultiColumnWeekMode = True
            frmInbox.CalendarControl.ViewType = xtpCalendarWeekView

        Case ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_MONTH_LOW, _
             ID_GROUP_ARRANGE_MONTH_MEDIUM, ID_GROUP_ARRANGE_MONTH_HIGH:
            frmInbox.CalendarControl.ViewType = xtpCalendarMonthView
            
        Case ID_CALENDAREVENT_OPEN:
            frmInbox.mnuOpenEvent
            
        Case ID_CALENDAREVENT_DELETE:
            frmInbox.mnuDeleteEvent
            
        Case ID_CALENDAREVENT_NEW, ID_GROUP_NEW_APPOINTMENT:
            'falta### frmEditEvent.AllDayOverride = False
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_MEETING:
            'falta### frmEditEvent.AllDayOverride = False
            'falta### frmEditEvent.chkMeeting.Value = 1
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_GROUP_NEW_ALLDAY:
            'falta### frmEditEvent.AllDayOverride = True
            frmInbox.mnuNewEvent
            frmInbox.CalendarControl.Options.DayViewCurrentTimeMarkVisible = True
            
        Case ID_CALENDAREVENT_CHANGE_TIMEZONE:
            frmInbox.mnuChangeTimeZone
            
        Case ID_CALENDAREVENT_60:
            frmInbox.mnuTimeScale 60
            
        Case ID_CALENDAREVENT_30:
            frmInbox.mnuTimeScale 30
            
        Case ID_CALENDAREVENT_15:
            frmInbox.mnuTimeScale 15
            
        Case ID_CALENDAREVENT_10:
            frmInbox.mnuTimeScale 10
            
        Case ID_CALENDAREVENT_5:
            frmInbox.mnuTimeScale 5
            
        Case Else
            AbiertoFormulario = True
            AbrirFormularios Control.Id
            
    End Select
    
    
    If AbiertoFormulario Then
        AbiertoFormulario = False
        'mOTIVO... no lo se
        'Pero si lo vamos cambiando funciona
        If Me.DockingPaneManager.Panes(1).Enabled = 3 Then
            Me.DockingPaneManager.Panes(1).Enabled = 3
            Me.DockingPaneManager.Panes(2).Enabled = 3

            frmPaneCalendar.DatePicker.Enabled = True
            
            DockingPaneManager.RedrawPanes
            
            
        Else
            Me.DockingPaneManager.Panes(1).Enabled = 3
            Me.DockingPaneManager.Panes(2).Enabled = 3
             
        End If
        DockingPaneManager.NormalizeSplitters

    End If
End Sub



Private Sub CommandBars_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
        Dim Control As CommandBarControl, ControlItem As CommandBarControl
        
        If TypeOf CommandBar Is RibbonBackstageView Then
            Debug.Print "RibbonBackstageView"
        End If
        
        Set Control = CommandBar.FindControl(, IDS_ARRANGE_BY)
        If Not Control Is Nothing Then
            Dim Index As Long
            Index = Control.Index
            Control.visible = False
            
            Do While Index + 1 <= CommandBar.Controls.Count
                Set ControlItem = CommandBar.Controls.Item(Index + 1)
                If ControlItem.Id = IDS_ARRANGE_BY Then
                    ControlItem.Delete
                Else
                    Exit Do
                End If
            Loop
            
'            Dim CurrentColumn As ReportColumn
'            For Each CurrentColumn In frmInbox. wndReportControl.Columns
'                Set ControlItem = CommandBar.Controls.Add(xtpControlButton, ID_REPORTCONTROL_COLUMN_ARRANGE_BY, CurrentColumn.Caption)
'                ControlItem.Parameter = CurrentColumn.ItemIndex
'                If Not frmInbox. wndReportControl.SortOrder.IndexOf(CurrentColumn) = -1 Then
'                    ControlItem.Checked = True
'                End If
'                If Not CurrentColumn.Visible Then
'                    ControlItem.Visible = False
'                End If
'            Next
        
        End If
End Sub

Private Sub CommandBars_SpecialColorChanged()
    Me.BackColor = CommandBars.GetSpecialColor(XPCOLOR_SPLITTER_FACE)
End Sub

Private Sub CommandBars_ToolBarVisibleChanged(ByVal ToolBar As XtremeCommandBars.ICommandBar)
     Debug.Print ToolBar.BarID
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
        
    On Error Resume Next
    
    
    
    Select Case Control.Id
        Case ID_VIEW_STATUSBAR:
            'Control.Checked = CommandBars.StatusBar.Visible
            
        Case ID_GROUP_ARRANGE_WORK_WEEK:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWorkWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_WEEK:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarWeekView, True, False)
            
        Case ID_GROUP_ARRANGE_MONTH:
            'Control.Checked = IIf(frmInbox.CalendarControl.ViewType = xtpCalendarMonthView, True, False)
        
        Case ID_OPTIONS_ANIMATION:
            'Control.Checked = CommandBars.ActiveMenuBar.EnableAnimation
            
        Case ID_OPTIONS_FONT_SYSTEM, ID_OPTIONS_FONT_NORMAL, ID_OPTIONS_FONT_LARGE, ID_OPTIONS_FONT_EXTRALARGE
             '   Dim newFontHeight As Integer
             '   newFontHeight = FontSizes(Control.Id - ID_OPTIONS_FONT_SYSTEM)
             '   Control.Checked = IIf(RibbonBar.FontHeight = newFontHeight, True, False)
                
        Case ID_OPTIONS_FONT_AUTORESIZEICONS
              '  Control.Checked = CommandBars.PaintManager.AutoResizeIcons

        Case ID_RIBBON_EXPAND:
            'Control.Visible = RibbonBar.Minimized
            
        Case ID_RIBBON_MINIMIZE:
            'Control.Visible = Not RibbonBar.Minimized
    End Select
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub DockingPaneManager_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, ByVal Container As XtremeDockingPane.IPaneActionContainer, Cancel As Boolean)
    If (Action = PaneActionSplitterResized) Then
        DockingPaneManager.RecalcLayout
        
        ' Save MRUShortcutBarWidth
        If (frmShortBar.ScaleWidth > MinimizedShortcutBarWidth And Container.Container.Type = PaneTypeSplitterContainer) Then
            Debug.Print frmShortBar.ScaleWidth
            MRUShortcutBarWidth = frmShortBar.ScaleWidth
        End If
    Else
        If (Action = PaneActionSplitterResized) Then Debug.Print "Resizing "
    End If
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Tag = PANE_SHORTCUTBAR Then
        Item.Handle = frmShortBar.hwnd
    ElseIf Item.Tag = PANE_REPORT_CONTROL Then
        Item.Handle = frmInbox.hwnd
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    End If
End Sub


Private Sub CargaDatosMenusDemas()
Dim AntiguoTab As Integer
    
    
    Screen.MousePointer = vbHourglass
    AntiguoTab = -1
    If RibbonSeHaCreado Then
        If Not RibbonBar.SelectedTab Is Nothing Then AntiguoTab = RibbonBar.SelectedTab.Id
    End If
    CreateRibbon
    Screen.MousePointer = vbHourglass
    CreateBackstage
    Screen.MousePointer = vbHourglass
    CreateRibbonOptions
    
    'vEmpresa.TieneContabilidad = False
    '??????
    '0=solo contabilidad / 1=todo / 2=solo tesoreria
    Screen.MousePointer = vbHourglass
    CargaMenu AntiguoTab
    CreateStatusBar
    Screen.MousePointer = vbHourglass
    PonerCaption
    CreateCalendarTabOriginal
    RibbonSeHaCreado = True
End Sub

Public Sub CambiarEmpresa(QueEmpresa As Integer)
Dim cur As Integer
    cur = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    Me.Hide
    CambiarEmpresa2 QueEmpresa
    Me.Show
    Screen.MousePointer = cur
    
End Sub

Public Sub CambiarEmpresa2(QueEmpresa As Integer)
Dim RB As RibbonBar

    CadenaDesdeOtroForm = vUsu.Login & "|" & vEmpresa.codempre & "|"
        
    Set vUsu = New USUARIO
    vUsu.Leer RecuperaValor(CadenaDesdeOtroForm, 1)
    
    If QueEmpresa = 0 Then
        vUsu.CadenaConexion = "aritaxi"
    Else
        vUsu.CadenaConexion = "aritaxi" & QueEmpresa
    End If
    
    AbrirConexion
    
    Set vEmpresa = New Cempresa
    Set vParam = New Cparametros
    '[Monica]11/01/2018: añadido el tema de empresas para cordoba
    Set vParamAplic = New CParamAplic
    
    'NO DEBERIAN DAR ERROR
    vEmpresa.LeerDatos
    vParam.Leer
    
    '[Monica]11/01/2018: añadido el tema de empresas para cordoba
    vParamAplic.Leer
    If AbrirConexionConta(False) = False Then
        MsgBox "La aplicación no puede continuar sin acceso a los datos. ", vbCritical
        End
    End If
    ' hasta aqui

    'Carga los Niveles de cuentas de Contabilidad de la empresa y las fechasINICIO y FIN
    LeerNivelesEmpresa
    
    PonerCaption
    
    Screen.MousePointer = vbHourglass
    
    CargaDatosMenusDemas
    
    frmPaneContacts.SeleccionarNodoEmpresa vEmpresa.codempre
    pageBackstageHelp.Label9.Caption = vEmpresa.nomempre
    pageBackstageHelp.tabPage(0).visible = False
    pageBackstageHelp.tabPage(1).visible = False
    frmInbox.OpenProvider
    
    Set RB = RibbonBar
    RB.Minimized = False
    RB.RedrawBar
   
   
  
   
'    vControl.UltEmpre = vUsu.CadenaConexion
'
'    vControl.Grabar
'
   '[Monica]23/01/2018
   'If vParamAplic.ContabilidadNueva And (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar
    If (vUsu.Nivel = 0 Or vUsu.Nivel = 1) Then FrasPendientesContabilizar

    
    
    Screen.MousePointer = vbDefault
End Sub



Private Sub Form_Load()
   
    'Cargamos librerias de icinos de los forms
    frmIdentifica.pLabel "Carga DLL"
    
    CargaIconosDlls
   
    
   
    CommandBarsGlobalSettings.App = App
            
    frmIdentifica.pLabel "Leyendo menus usuario"
    CargaDatosMenusDemas
    
    ShowEventInPane = False
       
    FontSizes(0) = 0
    FontSizes(1) = 11
    FontSizes(2) = 13
    FontSizes(3) = 16
               
    DockingPaneManager.SetCommandBars Me.CommandBars
              
    Set frmShortBar = New frmShortcutBar2
    Set frmInbox = New frmInbox
        
    Dim A As Pane, b As Pane, C As Pane, D As Pane
    
    frmIdentifica.pLabel "Creando paneles"
    Set A = DockingPaneManager.CreatePane(PANE_SHORTCUTBAR, 170, 120, DockLeftOf, Nothing)
    A.Tag = PANE_SHORTCUTBAR
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    
    Set b = DockingPaneManager.CreatePane(PANE_REPORT_CONTROL, 700, 400, DockRightOf, A)
    b.Tag = PANE_REPORT_CONTROL
   
    DockingPaneManager.Options.HideClient = True
    PonerTabPorDefecto -1
    
    Set CommandBars.Icons = CommandBarsGlobalSettings.Icons
    LoadIcons
    
    DockingPaneManager.RecalcLayout
    MRUShortcutBarWidth = frmShortBar.ScaleWidth
   
   
    'En funcion
    ' ID_OPTIONS_STYLEBLUE2010  ID_OPTIONS_STYLESILVER2010    ID_OPTIONS_STYLEBLACK2010
    frmIdentifica.pLabel "Carga skin"
    Screen.MousePointer = vbHourglass
    If vUsu.Skin = 3 Then
        Cad = ID_OPTIONS_STYLEBLACK2010
    Else
        If vUsu.Skin = 2 Then
            Cad = ID_OPTIONS_STYLESILVER2010
        Else
            Cad = ID_OPTIONS_STYLEBLUE2010
        End If
    End If
    CommandBars.FindControl(, Cad, , True).Execute
    
    PrimeraVez = True

    
End Sub


Private Sub CargaIconosDlls()
Dim TamanyoImgComun As Integer

'    ImageList1.ImageHeight = 48
'    ImageList1.ImageWidth = 48
'    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 48
'
'
'    ImageList2.ImageHeight = 16
'    ImageList2.ImageWidth = 16
'    GetIconsFromLibrary App.Path & "\styles\icoconppal.dll", 1, 16
'
'    ImageListPPal48.ImageHeight = 48
'    ImageListPPal48.ImageWidth = 48
'    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 8, 48
'
'
'    ImageListPpal16.ImageHeight = 16
'    ImageListPpal16.ImageWidth = 16
'    GetIconsFromLibrary App.Path & "\styles\icoconppal2.dll", 9, 16
'
'
'    imgListComun.ListImages.Clear
'    imgListComun_BN.ListImages.Clear
'    imgListComun_OM.ListImages.Clear
'
'        TamanyoImgComun = 24
'
'        imgListComun.ImageHeight = TamanyoImgComun
'        imgListComun.ImageWidth = TamanyoImgComun
'        GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 2, TamanyoImgComun  'antes icolistcon
'
'
'
'        '++
'        imgListComun_BN.ImageHeight = TamanyoImgComun
'        imgListComun_BN.ImageWidth = TamanyoImgComun
'        GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 3, TamanyoImgComun
'
'        imgListComun_OM.ImageHeight = TamanyoImgComun
'        imgListComun_OM.ImageWidth = TamanyoImgComun
'        GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 4, TamanyoImgComun
'
'
'    imgListComun16.ImageHeight = 16
'    imgListComun16.ImageWidth = 16
'    GetIconsFromLibrary App.Path & "\styles\iconosconta.dll", 5, 16
'
'    GetIconsFromLibrary App.Path & "\styles\iconosconta_BN.dll", 6, 16
'    GetIconsFromLibrary App.Path & "\styles\iconosconta_OM.dll", 7, 16


'????
    imgListComun1.ListImages.Clear
    imgListComun_BN.ListImages.Clear
    imgListComun_OM.ListImages.Clear
    
    TamanyoImgComun = 24
    
    imgListComun1.ImageHeight = TamanyoImgComun
    imgListComun1.ImageWidth = TamanyoImgComun
    GetIconsFromLibrary App.Path & "\styles\iconos.dll", 2, TamanyoImgComun  'antes icolistcon

    
    imgListComun_BN.ImageHeight = TamanyoImgComun
    imgListComun_BN.ImageWidth = TamanyoImgComun
    GetIconsFromLibrary App.Path & "\styles\iconos_BN.dll", 3, TamanyoImgComun
  
    imgListComun_OM.ImageHeight = TamanyoImgComun
    imgListComun_OM.ImageWidth = TamanyoImgComun
    GetIconsFromLibrary App.Path & "\styles\iconos_OM.dll", 4, TamanyoImgComun
    
    imgListComun16.ImageHeight = 16
    imgListComun16.ImageWidth = 16
    GetIconsFromLibrary App.Path & "\styles\iconos.dll", 5, 16
    
    GetIconsFromLibrary App.Path & "\styles\iconos_BN.dll", 6, 16
    GetIconsFromLibrary App.Path & "\styles\iconos_OM.dll", 7, 16



'????


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



Public Sub ExpandButtonClicked()
   
    
    
    Dim A As Pane
    Set A = DockingPaneManager.FindPane(PANE_SHORTCUTBAR)
    
    Dim ShortcutBarMinimized As Boolean
    ShortcutBarMinimized = frmShortBar.ScaleWidth <= MinimizedShortcutBarWidth
    
    Dim NewWidth As Long
    If (ShortcutBarMinimized) Then
        NewWidth = MRUShortcutBarWidth
    Else
        NewWidth = MinimizedShortcutBarWidth
        frmShortBar.wndShortcutBar.PopupWidth = MRUShortcutBarWidth
    End If
        
    
    ' Set Size of Pane
    A.MinTrackSize.Width = NewWidth
    A.MaxTrackSize.Width = NewWidth
        
    DockingPaneManager.RecalcLayout
    DockingPaneManager.NormalizeSplitters
    DockingPaneManager.RedrawPanes
    
    ' Restore Constraints
    A.MinTrackSize.Width = MinimizedShortcutBarWidth
    A.MaxTrackSize.Width = 32000
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If Not (pageBackstageInfo Is Nothing) Then Unload pageBackstageInfo
    If Not (pageBackstageHelp Is Nothing) Then Unload pageBackstageHelp
    'If Not (pageBackstageSend Is Nothing) Then Unload pageBackstageSend
    
    'close all sub forms
    On Error Resume Next
    Dim i As Long
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
    GuardarDatosUltimaTab
  
    AccionesCierre
  
End Sub


Private Sub GuardarDatosUltimaTab()
    i = RibbonBar.SelectedTab.Id
    If i = ID_TAB_CALENDAR_HOME Then Exit Sub 'no guardo este tab
    If i <> vUsu.TabPorDefecto Then
        vUsu.TabPorDefecto = i
        vUsu.GuardarTabPorDefecto
    End If
End Sub


Private Sub AccionesCierre()

    NumeroEmpresaMemorizar False

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


Private Sub NumeroEmpresaMemorizar(Leer As Boolean)
Dim NF As Integer
Dim C1 As String
Dim Cad As String

On Error GoTo ENumeroEmpresaMemorizar

    Cad = App.Path & "\ultempre.dat"
    If Leer Then
        If Dir(Cad) <> "" Then
            NF = FreeFile
            Open Cad For Input As #NF
            Line Input #NF, Cad
            Close #NF
            Cad = Trim(Cad)
            If Cad <> "" Then
                vUsu.CadenaConexion = RecuperaValor(Cad, 2)
            End If
        End If
    Else 'Escribir
        NF = FreeFile
        Open Cad For Output As #NF
        Cad = "NO ncesito|" & vUsu.CadenaConexion & "|"
        Print #NF, Cad
        Close #NF
    End If
    
ENumeroEmpresaMemorizar:
    Err.Clear
End Sub


Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, Caption)
    
    Control.BeginGroup = BeginGroup
    Control.DescriptionText = DescriptionText
    Control.Style = ButtonStyle
    Control.Category = Category
    
    Set AddButton = Control
    
End Function

Private Sub CommandBars_Resize()
    
    On Error Resume Next
    
    Dim Left As Long
    Dim top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, top, Right, Bottom
    
End Sub

Private Sub LoadIcons()
    CommandBars.Icons.RemoveAll
    SuiteControlsGlobalSettings.Icons.RemoveAll
    ReportControlGlobalSettings.Icons.RemoveAll

    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\help.png", ID_APP_ABOUT, xtpImageNormal
        
    'Para que no carge imagen de ratios y graficas y punteo, no lo pongo aqui ya que los cargo "pequeños"
    '
      
    'ICONOS PEQUEÑOS
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_Clientes, ID_ClientesInactivos, ID_TiposCartas, ID_MotivosBajaEquipos, ID_TiposAveria, ID_FormasPago), xtpImageNormal
        
    
    'Pequeños
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(ID_CartasClientes, ID_Marcas, ID_TiposUnidad, ID_ClientesAgrup, ID_SituEspe, ID_Incidencias, ID_Tarjetas, ID_BancosPropio, ID_AltasClientes, ID_InfClientes), xtpImageNormal
        
    'Pequeños diario
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_AlmacenesPropios, ID_FamiliasArticulos, ID_ArticulosInactivos, ID_MovimientosAlmacen, ID_HistoricoMovimientosAlmacen, ID_MovimientosArticulos), xtpImageNormal
      
         '      ID_AsientosPredefinidos
     
    'Deberiamos cargar un array con unos(1) de longitud 143
    ' y en funcion del valor del campo imagen en el punto de menu correspondiente
    ' lo pondremos en el array.
    ' Ejemplo    303 Extractos  Campo imagen: 87
    ' quiere decir que en el campo 87 del array sustituieremos el 1 por el 303


'
    Dim T() As Variant
    'Cad linea son 15
    T = Array(1, 1, ID_TiposArticulos, 1, ID_PreciosProv, 1, 1, ID_ContaFacturas, ID_FrasRectificativas, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, ID_RetenSocios, ID_ContabilFras, ID_HcoInventario, ID_Reparaciones, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, ID_ValStocksInven, ID_Articulos, ID_Socios, ID_Empresa, ID_ParametrosContabilidad, 1, ID_Contadores, 1, ID_ControlRep, 1, 1, 1, 1, 1, _
        ID_ArticulosComponentes, 1, 1, ID_ServSocios, ID_RecepFacturas, ID_Informes, 1, ID_Usuarios, ID_HcoLlamadas, ID_TraspasoTaxitronic, 1, 1, ID_SelImpresora, ID_ConfigurarBalances, 1, _
        1, ID_ActDiferencias, ID_EtiProveedores, 1, 1, 1, ID_ReimprFras, ID_HcoFras, ID_Liquidacion, ID_Choferes, 1, 1, ID_HcoFacturas, 1, 1, _
        ID_ProvVarios, 1, ID_FacturacionSocios, ID_HistoricoUves, 1, ID_Clientes, 1, 1, 1, 1, ID_ContabFacturas, 1, 1, ID_Proveedores, 1, _
        ID_Direcciones, ID_CartasProv, 1, ID_InfProveedores, 1, 1, ID_PedidosProv, 1, 1, ID_AlbAnuladosPro, ID_AlbProveedor, ID_HcoAlbxFra, 1, ID_ComprasProveedor, 1, _
        ID_TomaInventario, ID_ValoracionStocks, 1, ID_DtosProv, ID_VtasFamArt, ID_VtasMeses, ID_ContabFras, ID_Albaranes, 1, ID_FacturarClientes, ID_StocksMaxMin, 1, 1, 1, 1, _
        ID_AlbxArt, ID_StocksFecha, 1, 1, ID_EntradaExisReal, 1, ID_ComprasFamxArt, ID_PedidosAnulados, ID_ListadoMovimientos, 1, 1, 1, 1, ID_AlbAnulados, 1, _
        ID_ListadoDiferencias, 1, ID_InfAlbxProv, ID_DetalleFra, 1)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", T, xtpImageNormal

    T = Array(ID_FrasRectifCli, ID_FrasRectifSocios, ID_HcoFrasClientes, ID_HcoFacturasSocios, ID_ReimprFrasSocios, ID_ContabFrasSocios, _
        1, 1, 1, 1)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal

    'pequeños
    T = Array(ID_EtiquetasClientes, ID_VentasporCliente, ID_DetalleFacturacion, ID_AgentesCom, ID_Actividades, ID_ServiciosAbonados, ID_FacturacionClientes, ID_FactuVarClientes, ID_ReimprimirFras, ID_EtiquetasSocios)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal



    'pequeños
    T = Array(ID_VtasSocios, ID_ServiciosSocios, ID_LiquidacionesSocios, ID_Aportaciones, 1, 1, 1, 1, 1, 1)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal



    'cuotas
    T = Array(1, ID_Facturacion, ID_PrevFacturacCuotas, ID_HcoFrasCuotas, ID_MtoAlbaranes, ID_ContabFrasCuotas, 1, ID_ReimprFrasCuotas, ID_CartasSocios, ID_FrasRectific)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal






    'acciones comerciales
    T = Array(1, 1, 1, ID_TiposAcciones, 1, 1, _
        1, ID_GenerarAcciones, 1, ID_AccionesComer)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal


    ' UTILIDADES
    T = Array(ID_FrasPdtesContabilizar, ID_Llamadas, ID_EliminarArticulos, ID_RevisarCaracteresEsp, ID_CopiaSeguridadLocal, ID_ErroresNrosFrasCliente, _
        ID_BorreFrasMovimientos, ID_ConceptosLlamadas, ID_FacturacionElectronica, ID_BorreFrasMovimientos)
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal

    T = Array(1, ID_UsuariosActivos, ID_Avisos, ID_AccionesRealizadas, ID_ConexionesActivas, 1, 1, 1, 1, 1)

    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", T, xtpImageNormal


    'Este de abjo funciona correctamente.
    'NO tocar. Es por si falla volver a empezar
'    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlook2013L_32x32.bmp", _
'            Array(ID_CarteradeCobros, ID_InformeCobrosPendientes, ID_RealizarCobro, ID_Compensarcliente, 1, ID_BalancePresupuestario, 1, _
'            ID_CentrosdeCoste, 1, 1, ID_Presupuestos, ID_Remesas, ID_Detalledeexplotación, ID_CarteradePagos, ID_CuentadeExplotaciónAnalítica, ID_ExtractosporCentrodeCoste, _
'            ID_Asientos, ID_Extractos, ID_Punteo, 1, ID_CuentadeExplotación, ID_Totalesporconcepto, ID_BalancedeSituación, ID_PérdidasyGanancias, _
'            ID_SumasySaldos, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            ID_Empresa, ID_ParametrosContabilidad, ID_Contadores, ID_Usuarios, 1, ID_Informes, ID_Nuevaempresa, ID_ConfigurarBalances, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            ID_FacturasEmitidas, ID_LibroFacturasEmitidas, ID_FacturasRecibidas, ID_LibroFacturasRecibidas, 1, 1, 1, 1, 1, ID_Elementos, ID_GenerarAmortización, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
'            1, ID_PlanContable, ID_TiposdeDiario, ID_Conceptos, ID_TiposdeIVA, ID_TiposdePago, ID_Bancos, ID_FormasdePago, _
'            ID_BicSwift, ID_Agentes, ID_AsientosPredefinidos, ID_ModelosdeCartas, _
'            ID_Renumeracióndeasientos, ID_CierredeEjercicio, ID_Deshacercierre, 1, 1, 1, 1, 1, 1, ID_DiarioOficial, _
'            ID_PresentaciónTelemáticadeLibros, ID_Traspasodecuentasenapuntes, ID_Renumerarregistrosproveedor, 1, ID_TraspasocodigosdeIVA), xtpImageNormal
'
    
    'Presupuiestaria y analitaica cargadas arriba en pequeño
    '---------------------------------------------------------
    '
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_Vehiculos, ID_Trabajadores _
                  ), xtpImageNormal
    
    'Pequeños
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(ID_PrevFacturacion, ID_FacturacionAlb, ID_FacturasRect, ID_HcoAlbFra, ID_ReimprirFras, ID_NrosSerie, _
            ID_MotivosPdteRepara, ID_ServAsistenciaTecnica, _
            ID_InfPdteFacturar, ID_MatPdteRecibir _
            ), xtpImageNormal
    
    
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_TrabajosRealizados, ID_InfReparacionesDia, ID_Informeporcuenta, ID_SituaciónTesoreria, ID_InformeporNIF, _
                   ID_GrarFrasCuotas), xtpImageNormal
  
        
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------
    '------------------------------------------------------------------------------------------------------------------------
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookcalicons.png", _
            Array(ID_GROUP_NEW_APPOINTMENT, ID_GROUP_NEW_MEETING, ID_GROUP_NEW_ITEMS, ID_GROUP_GOTO_TODAY, _
            ID_GROUP_GOTO_NEXT7DAYS, ID_GROUP_ARRANGE_DAY, ID_GROUP_ARRANGE_WORK_WEEK, ID_GROUP_ARRANGE_WEEK, _
            ID_GROUP_ARRANGE_MONTH, ID_GROUP_ARRANGE_SCHEDULE_VIEW, ID_GROUP_MANAGE_CALENDARS_OPEN, ID_GROUP_MANAGE_CALENDARS_GROUPS, _
            ID_GROUP_SHARE_EMAIL, ID_GROUP_SHARE_SHARE, ID_GROUP_SHARE_PUBLISH, ID_GROUP_SHARE_PERMISSIONS), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\RibbonMinimize.png", _
            Array(ID_RIBBON_MINIMIZE, ID_RIBBON_EXPAND), xtpImageNormal
            
    CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\Search.png", _
            ID_SEARCH_ICON, xtpImageNormal
            
     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonslarge.png", _
            Array(ID_GROUP_MAIL_NEW_NEW, ID_GROUP_MAIL_NEW_NEW_ITEMS, ID_GROUP_MAIL_DELETE_DELETE, ID_GROUP_MAIL_RESPOND_REPLY, _
            ID_GROUP_MAIL_RESPOND_REPLY_ALL, ID_GROUP_MAIL_RESPOND_FORWARD, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
            
     CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\reporticonssmall.png", _
            Array(ID_GROUP_MAIL_DELETE_CLEANUP, ID_GROUP_MAIL_DELETE_JUNK, ID_GROUP_MAIL_RESPOND_MEETING, ID_GROUP_MAIL_RESPOND_IM, _
            ID_GROUP_MAIL_RESPOND_MORE, ID_GROUP_MAIL_TAGS_UNREAD, ID_GROUP_MAIL_TAGS_CATEGORIZE, ID_GROUP_MAIL_TAGS_FOLLOWUP, ID_GROUP_MAIL_FIND_ADDRESSBOOK, _
            ID_GROUP_MAIL_FIND_FILTER, ID_GROUP_MAIL_MOVE_MOVE, ID_GROUP_MAIL_MOVE_ONENOTE), xtpImageNormal
    
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\outlookpane.png", _
            Array(ID_SWITCH_NORMAL, ID_SWITCH_CALENAR_AND_TASK, ID_SWITCH_CALENDAR, ID_SWITCH_CLASSIC, ID_SWITCH_READING), xtpImageNormal
            
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_16x16.bmp", _
            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\mail_24x24.bmp", _
            Array(SHORTCUT_INBOX, SHORTCUT_CALENDAR, SHORTCUT_CONTACTS, SHORTCUT_TASKS, SHORTCUT_NOTES, _
            SHORTCUT_FOLDER_LIST, SHORTCUT_SHORTCUTS, SHORTCUT_JOURNAL, SHORTCUT_SHOW_MORE, SHORTCUT_SHOW_FEWER), xtpImageNormal
            
        CommandBars.Icons.LoadBitmap App.Path & "\styles\quickstepsgallery.png", _
            Array(ID_QUICKSTEP_REPLAY_DELETE, ID_QUICKSTEP_TO_MANAGER, ID_QUICKSTEP_MOVE_TO, ID_QUICKSTEP_CREATE_NEW, ID_QUICKSTEP_TEAM_EMAIL, ID_QUICKSTEP_DONE), xtpImageNormal
            
        ReportControlGlobalSettings.Icons.LoadBitmap App.Path & "\styles\bmreport.bmp", _
        Array(COLUMN_MAIL_ICON, COLUMN_IMPORTANCE_ICON, COLUMN_CHECK_ICON, RECORD_UNREAD_MAIL_ICON, RECORD_READ_MAIL_ICON, _
            RECORD_REPLIED_ICON, RECORD_IMPORTANCE_HIGH_ICON, COLUMN_ATTACHMENT_ICON, COLUMN_ATTACHMENT_NORMAL_ICON, _
            RECORD_IMPORTANCE_LOW_ICON), xtpImageNormal
            
            
        CommandBarsGlobalSettings.Icons.LoadBitmap App.Path & "\styles\car_compact_grey.bmp", ID_Vehiculos, xtpImageNormal
            
            
        Dim i As Integer
        For i = 1 To 17
            SuiteControlsGlobalSettings.Icons.LoadIcon App.Path & "\styles\TreeView\icon" & i & ".ico", i, xtpImageNormal
        Next i
End Sub

Private Sub SaveRibbonBarToXML()
    Dim Px As PropExchange
    Set Px = XtremeCommandBars.CreatePropExchange()
    
    Px.CreateAsXML False, "Settings"
        
    Dim Options As StateOptions
    Set Options = CommandBars.CreateStateOptions()
    Options.SerializeControls = True
        
    CommandBars.DoPropExchange Px.GetSection("CommandBars"), Options
    
    Px.SaveToFile "C:\Layout.xml"
    
End Sub



Private Function CreateQuickStepGallery() As CommandBarGalleryItems

    Dim GalleryItems As CommandBarGalleryItems
    Set GalleryItems = CommandBars.CreateGalleryItems(ID_GALLERY_QUICKSTEP)
        
    GalleryItems.ItemWidth = 120
    GalleryItems.ItemHeight = 20
            
    GalleryItems.AddItem ID_QUICKSTEP_MOVE_TO, "Move To: ?"
    GalleryItems.AddItem ID_QUICKSTEP_TO_MANAGER, "To Manager"
    GalleryItems.AddItem ID_QUICKSTEP_TEAM_EMAIL, "Team E-mail"
    GalleryItems.AddItem ID_QUICKSTEP_DONE, "Done"
    GalleryItems.AddItem ID_QUICKSTEP_REPLAY_DELETE, "Reply & Delete"
    GalleryItems.AddItem ID_QUICKSTEP_CREATE_NEW, "Create New"
        
    GalleryItems.Icons = CommandBarsGlobalSettings.Icons

    Set CreateQuickStepGallery = GalleryItems

End Function

Private Sub CommandBars_ControlNotify(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal Code As Long, ByVal NotifyData As Variant, Handled As Variant)
   
    If (Code = XTP_BS_TABCHANGED) Then

        
    End If
End Sub


Private Sub CreateBackstage()

    
    Dim RibbonBar As RibbonBar
    Set RibbonBar = CommandBars.ActiveMenuBar
    
    Dim BackstageView As RibbonBackstageView
    Set BackstageView = CommandBars.CreateCommandBar("CXTPRibbonBackstageView")
    
    BackstageView.SetTheme xtpThemeRibbon


    CommandBars.Icons.LoadBitmap App.Path & "\styles\BackstageIcons.png", _
    Array(1, 1, 1002, 1, 1, ID_APP_EXIT), xtpImageNormal

    Set RibbonBar.AddSystemButton.CommandBar = BackstageView
    
    'BackstageView.AddCommand ID_FILE_SAVE, "Cambiar empresa"
    'BackstageView.AddCommand ID_FILE_SAVE_AS, "Personalizar"
    'BackstageView.AddCommand ID_FILE_OPEN, "Open"
    'BackstageView.AddCommand ID_FILE_CLOSE, "Close"
    
    'If (pageBackstageInfo Is Nothing) Then Set pageBackstageInfo = New pageBackstageInfo
    'If (pageBackstageSend Is Nothing) Then Set pageBackstageSend = New pageBackstageSend
    If (pageBackstageHelp Is Nothing) Then Set pageBackstageHelp = New pageBackstageHelp
    
    Dim ControlInfo As RibbonBackstageTab
    Set ControlInfo = BackstageView.AddTab(1000, "Info", pageBackstageHelp.hwnd)
    
    'BackstageView.AddTab 1002, "Empresas", pageBackstageSend.hwnd

    ' Los menus de informacion...
    'BackstageView.AddTab 1001, "Acerca de", pageBackstageInfo.hwnd
    
    
    
    
    
    
    
    
    
    
    'BackstageView.AddCommand ID_FILE_OPTIONS, "Options"
    BackstageView.AddCommand ID_APP_EXIT, "Salir"
    
    ControlInfo.DefaultItem = True
    

End Sub




Private Sub CreateCalendarTabOriginal()

    Dim TabCalendarHome As RibbonTab
    Dim GroupNew As RibbonGroup, GroupGoTo As RibbonGroup, GroupArrange As RibbonGroup

    
    Dim Control As CommandBarControl
    Dim ControlNew_NewItems As CommandBarPopup
    Dim ControlArrange_Month As CommandBarPopup
    Dim ControlManage_Open As CommandBarPopup
    Dim ControlManage_Groups As CommandBarPopup
    Dim ControlShare_Publish As CommandBarPopup
           
    Dim PopupBar As CommandBar
    
    Set TabCalendarHome = RibbonBar.InsertTab(14, "Agenda")
    TabCalendarHome.Id = ID_TAB_CALENDAR_HOME
 
    Set GroupNew = TabCalendarHome.Groups.AddGroup("&Nueva", ID_GROUP_NEW)
        
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "&Evento")
    Control.Enabled = False
    Set Control = GroupNew.Add(xtpControlButton, ID_GROUP_NEW_MEETING, "&Cita")
    Control.Enabled = True
    
    '------------------------------------
    'Set ControlNew_NewItems = GroupNew.Add(xtpControlButtonPopup, ID_GROUP_NEW_ITEMS, "New &Items")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_APPOINTMENT, "Evento")
    '    Set Control = ControlNew_NewItems.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_NEW_ALLDAY, "E&vento todo el dia")
    '    Control.BeginGroup = True
    'ControlNew_NewItems.KeyboardTip = "V"
    
    Set GroupGoTo = TabCalendarHome.Groups.AddGroup("I&r a", ID_GROUP_GOTO)
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_TODAY, "&Hoy")
    Set Control = GroupGoTo.Add(xtpControlButton, ID_GROUP_GOTO_NEXT7DAYS, "Próximos &7 dias ")
    GroupGoTo.ShowOptionButton = True
    GroupGoTo.ControlGroupOption.Caption = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.ToolTipText = "Ir a (Ctrl+G)"
    GroupGoTo.ControlGroupOption.DescriptionText = "Ir a fecha especificada."
    
    Set GroupArrange = TabCalendarHome.Groups.AddGroup("Vista", ID_GROUP_ARRANGE2)
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_DAY, "&Dia vista")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WORK_WEEK, "Samana &trabajo")
    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_WEEK, "Sema&na vista")
    Set ControlArrange_Month = GroupArrange.Add(xtpControlSplitButtonPopup, ID_GROUP_ARRANGE_MONTH, "Mes")
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_LOW, "Ver detalle")
            Control.ToolTipText = "Muestra solo eventos todo el dia."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_MEDIUM, "Detalle &Medio")
            Control.ToolTipText = "Eventos todo el dia y si esta libre el dia o tiene eventos."
            Control.DescriptionText = Control.ToolTipText
            Set Control = ControlArrange_Month.CommandBar.Controls.Add(xtpControlButton, ID_GROUP_ARRANGE_MONTH_HIGH, "Detalle &Alto")
            Control.ToolTipText = "Muestra todo."
            Control.DescriptionText = Control.ToolTipText

'    Set Control = GroupArrange.Add(xtpControlButton, ID_GROUP_ARRANGE_SCHEDULE_VIEW, "Schedule View")
'    GroupArrange.ShowOptionButton = True
'    GroupArrange.ControlGroupOption.Caption = "Calendar Options"
'    GroupArrange.ControlGroupOption.ToolTipText = "Calendar Options"
'    GroupArrange.ControlGroupOption.DescriptionText = "Change the settings for calendars, meetings and time zones."
'
'
  
    
End Sub





Private Sub CreateRibbon()
    Dim RibbonBar As RibbonBar
    
    If RibbonSeHaCreado Then Exit Sub
        
    
    
    Set RibbonBar = CommandBars.AddRibbonBar("The Ribbon")
    RibbonBar.EnableDocking xtpFlagStretched
    
    RibbonBar.AllowQuickAccessCustomization = False
    RibbonBar.ShowQuickAccessBelowRibbon = False
    RibbonBar.ShowGripper = False
    
    RibbonBar.AllowMinimize = False
    RibbonBar.AddSystemButton
    
    RibbonBar.SystemButton.IconId = ID_SYSTEM_ICON
    RibbonBar.SystemButton.Caption = "&Menu"
    RibbonBar.SystemButton.Style = xtpButtonCaption
End Sub

Private Sub CreateRibbonOptions()

    CommandBars.EnableActions
    If RibbonSeHaCreado Then Exit Sub
    
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue", "Office 2010 Blue", "Office 2010 Blue", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver", "Office 2010 Silver", "Office 2010 Silver", "Themes"
    CommandBars.Actions.Add ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black", "Office 2010 Black", "Office 2010 Black", "Themes"

    Dim Control As CommandBarControl, ControlAbout As CommandBarControl
    Dim ControlPopup As CommandBarPopup, ControlOptions As CommandBarPopup
         
    Set ControlOptions = RibbonBar.Controls.Add(xtpControlPopup, 0, "Opciones")
    ControlOptions.Flags = xtpFlagRightAlign
    
    Set Control = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Styles")
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLUE2010, "Office 2010 Blue"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLESILVER2010, "Office 2010 Silver"
    Control.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_STYLEBLACK2010, "Office 2010 Black"
    
    Set ControlPopup = ControlOptions.CommandBar.Controls.Add(xtpControlPopup, 0, "Tamaño fuente", -1, False)
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_SYSTEM, "Sistema", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlRadioButton, ID_OPTIONS_FONT_NORMAL, "Normal", -1, False)
    Control.BeginGroup = True
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_LARGE, "Grande", -1, False
    ControlPopup.CommandBar.Controls.Add xtpControlRadioButton, ID_OPTIONS_FONT_EXTRALARGE, "Extra grande", -1, False
    Set Control = ControlPopup.CommandBar.Controls.Add(xtpControlButton, ID_OPTIONS_FONT_AUTORESIZEICONS, "Ajustar Icons", -1, False)
    Control.BeginGroup = True
    
    'ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_RTL, "Right To Left"
    ControlOptions.CommandBar.Controls.Add xtpControlButton, ID_OPTIONS_ANIMATION, "Animation   "
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_MINIMIZE, "Minimizar la barra", False, "Muestra solo los titulos del menu principal.")
    Control.Flags = xtpFlagRightAlign
    
    Set Control = AddButton(RibbonBar.Controls, xtpControlButton, ID_RIBBON_EXPAND, "Expandir la barra", False, "Muestra todos los elementos del menu.")
    Control.Flags = xtpFlagRightAlign
        
    Set ControlAbout = RibbonBar.Controls.Add(xtpControlButton, ID_APP_ABOUT, "&Acerca de")
    ControlAbout.Flags = xtpFlagRightAlign Or xtpFlagManualUpdate
    

        
End Sub








'*************************************************************************
'*************************************************************************
'*************************************************************************
'
'       CARGA menus en Ribbon
'
'




Public Sub CargaMenu(AntiguoTab As Integer)
Dim RN As ADODB.Recordset




    Set RN = New ADODB.Recordset
    Set Rn2 = New ADODB.Recordset
    On Error GoTo eCargaMenu
    

    If RibbonSeHaCreado Then RibbonBar.RemoveAllTabs
    
    Cad = "Select * from menus where aplicacion = 'aritaxi' and padre = 0 ORDER BY padre,orden "
    RN.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RN.EOF
    
        
        If Not BloqueaPuntoMenu(RN!Codigo, "aritaxi") Then
             Habilitado = True
             
             If Not MenuVisibleUsuario(DBLet(RN!Codigo), "aritaxi") Then
                 Habilitado = False
             Else
         
                 If (MenuVisibleUsuario(DBLet(RN!Padre), "aritaxi") And DBLet(RN!Padre) <> 0) Or DBLet(RN!Padre) = 0 Then
                     'OK todo habilitado
                 Else
                     Habilitado = False
                 End If
             End If
             
            
                
            If Habilitado Then
                
                Select Case RN!Codigo
                Case 1
                    '1   "CONFIGURACION"
                    CargaMenuConfiguracion RN!Codigo
                Case 2
                    '2   "ALMACEN"
                    CargaMenuAlmacen RN!Codigo
                Case 3
                    '3   "FACTURACION CLIENTES"
                    CargaMenuFacturaCliente RN!Codigo
                Case 4
                    '4   "FACTURACION SOCIOS"
                    CargaMenuFacturaSocio RN!Codigo
                Case 5
                    '5   "COMPRAS"
                    CargaMenuCompras RN!Codigo
                Case 6
                    '6   "PUBLICIDAD"
                    CargaMenuPublicidad RN!Codigo
                Case 7
                    '7   "CUOTAS"
                    CargaMenuCuotas RN!Codigo
                Case 8
                    '8   "REPARACIONES"
                    CargaMenuReparaciones RN!Codigo
                Case 9
                    '9   "CRM"
                     CargaMenuCRM RN!Codigo
                Case 10
                    '10  "UTILIDADES"
                    CargaMenuUtilidades RN!Codigo
'                Case Else
'                    MsgBox "Menu no tratado"
'                    End
                End Select
                
            End If
                                                 
        End If  'de habilitado el padre
    
        RN.MoveNext
    Wend
    RN.Close
                        
    PonerTabPorDefecto AntiguoTab
    
eCargaMenu:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    
    Set TabNuevo = Nothing
    Set GroupNew = Nothing
    Set Control = Nothing
    Set RN = Nothing
    Set Rn2 = Nothing
End Sub

Private Sub PonerTabPorDefecto(AntiguoTabSeleccionado As Integer)
Dim Anterior As Integer

    On Error Resume Next
    
    If AntiguoTabSeleccionado < 0 Then
        Anterior = vUsu.TabPorDefecto
    Else
        Anterior = AntiguoTabSeleccionado
    End If
    
    Cad = ""
    For i = 0 To RibbonBar.TabCount - 1
        J = RibbonBar.Tab(i).Id
        'Debug.Print J & " " & RibbonBar.Tab(i).Caption
        If J = Anterior Then
            
            RibbonBar.Tab(i).visible = True
            RibbonBar.Tab(i).Selected = True
            Set RibbonBar.SelectedTab = RibbonBar.Tab(i)
            Cad = "OK"
            Exit For
        End If
    Next
    If Cad = "" Then
        
        For J = RibbonBar.TabCount To 1 Step -1
            RibbonBar.Tab(J - 1).visible = True
            RibbonBar.Tab(J - 1).Selected = True
        Next J
    End If

    Err.Clear
End Sub

Private Sub CargaMenuConfiguracion(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Configuracion")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
        
        'color Categorias  eventos
'[Monica]24/01/2018: he quitado categorias calendario
'        If Not GroupNew Is Nothing Then
'            Set Control = GroupNew.Add(xtpControlButton, 199, "Categorias calendario")
'        End If
        Set GroupNew = Nothing
End Sub






Private Sub CargaMenuAlmacen(IdMenu As Integer)
Dim GrupGnral As RibbonGroup
Dim GrupMovim As RibbonGroup
Dim GrupConsu As RibbonGroup
Dim GrupInven As RibbonGroup
Dim SegundoGrupo As RibbonGroup

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Almacen")
        TabNuevo.Id = CLng(IdMenu)
        
        
        'En este llevaremos dos solapas, tesoreria y contabilidad (no le ponemos nombres)
        Cad = CStr(IdMenu * 100000)
        Set GrupGnral = TabNuevo.Groups.AddGroup("DATOS GENERALES", Cad & "0")
        Set GrupMovim = TabNuevo.Groups.AddGroup("MOVIMIENTOS ALMACEN", Cad & "1")
        Set GrupConsu = TabNuevo.Groups.AddGroup("CONSULTAS", Cad & "2")
        Set GrupInven = TabNuevo.Groups.AddGroup("INVENTARIO", Cad & "3")
        
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
             
                Select Case Rn2!Codigo
                    Case 201, 202, 203, 204, 205, 206
                        Set Control = GrupGnral.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 207, 208
                        Set Control = GrupMovim.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 209, 211, 212, 213, 214, 215, 216
                        Set Control = GrupConsu.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 217 To 222
                        Set Control = GrupInven.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                End Select
                
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close

         Set GroupNew = Nothing
End Sub


Private Sub CargaMenuFacturaCliente(IdMenu As Integer)
Dim GrupGral As RibbonGroup
Dim GrupVari As RibbonGroup
Dim GrupEsta As RibbonGroup
Dim GrupOtro As RibbonGroup


        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Facturación Clientes")
        TabNuevo.Id = CLng(IdMenu)
        
        Cad = CStr(IdMenu * 100000)
        Set GrupGral = TabNuevo.Groups.AddGroup("DATOS GENERALES", Cad & "0")
        Set GrupVari = TabNuevo.Groups.AddGroup("INFORMES VARIOS", Cad & "1")
        Set GrupOtro = TabNuevo.Groups.AddGroup("FACTURACION", Cad & "2")
        Set GrupEsta = TabNuevo.Groups.AddGroup("ESTADISTICA", Cad & "3")
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
                
                Select Case Rn2!Codigo
                    Case 301 To 310
                        Set Control = GrupGral.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 311 To 315
                        Set Control = GrupVari.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 316 To 324
                        Set Control = GrupOtro.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 325, 326
                        Set Control = GrupEsta.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                End Select
                
                
                Control.Enabled = Habilitado
              
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
    
        Set GrupGral = Nothing
        Set GrupVari = Nothing
        Set GrupEsta = Nothing
        Set GrupOtro = Nothing
End Sub

Private Sub CargaMenuFacturaSocio(IdMenu As Integer)
Dim GrupGral As RibbonGroup
Dim GrupVari As RibbonGroup
Dim GrupEsta As RibbonGroup
Dim GrupAlba As RibbonGroup
Dim GrupLiqu As RibbonGroup


        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Facturación Socios")
        TabNuevo.Id = CLng(IdMenu)
        
        Cad = CStr(IdMenu * 100000)
        Set GrupGral = TabNuevo.Groups.AddGroup("DATOS GENERALES", Cad & "0")
        Set GrupVari = TabNuevo.Groups.AddGroup("INFORMES VARIOS", Cad & "1")
        Set GrupAlba = TabNuevo.Groups.AddGroup("ALBARANES", Cad & "2")
        Set GrupLiqu = TabNuevo.Groups.AddGroup("LIQUIDACION", Cad & "3")
        Set GrupEsta = TabNuevo.Groups.AddGroup("ESTADISTICA", Cad & "4")
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
                
                Select Case Rn2!Codigo
                    Case 401 To 405
                        Set Control = GrupGral.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 406, 407, 408
                        Set Control = GrupVari.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 409 To 417
                        Set Control = GrupAlba.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 418 To 423
                        Set Control = GrupLiqu.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 424 To 429
                        Set Control = GrupEsta.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case Else
                        MsgBox "No se ha encontrado el código " & Rn2!Codigo & " " & Rn2!Descripcion, vbExclamation
                End Select
                
                
                Control.Enabled = Habilitado
              
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
    
        Set GrupGral = Nothing
        Set GrupVari = Nothing
        Set GrupEsta = Nothing
        Set GrupLiqu = Nothing
        Set GrupAlba = Nothing
        
End Sub

Private Sub CargaMenuCompras(IdMenu As Integer)
Dim GrupGral As RibbonGroup
Dim GrupVari As RibbonGroup
Dim GrupPrec As RibbonGroup
Dim GrupPedi As RibbonGroup
Dim GrupAlba As RibbonGroup
Dim GrupEsta As RibbonGroup


        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Compras")
        TabNuevo.Id = CLng(IdMenu)
        
        Cad = CStr(IdMenu * 100000)
        Set GrupGral = TabNuevo.Groups.AddGroup("DATOS GENERALES", Cad & "0")
        Set GrupVari = TabNuevo.Groups.AddGroup("INFORMES VARIOS", Cad & "1")
        Set GrupPrec = TabNuevo.Groups.AddGroup("PRECIOS", Cad & "2")
        Set GrupPedi = TabNuevo.Groups.AddGroup("PEDIDOS", Cad & "3")
        Set GrupAlba = TabNuevo.Groups.AddGroup("ALBARANES", Cad & "4")
        Set GrupEsta = TabNuevo.Groups.AddGroup("ESTADISTICA", Cad & "5")
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
        
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
                
                Select Case Rn2!Codigo
                    Case 501 To 503
                        Set Control = GrupGral.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 504 To 506
                        Set Control = GrupVari.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 507, 508
                        Set Control = GrupPrec.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 509 To 511
                        Set Control = GrupPedi.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 512 To 517
                        Set Control = GrupAlba.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    Case 518 To 520
                        Set Control = GrupEsta.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                End Select
                
                
                Control.Enabled = Habilitado
              
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
    
        Set GrupGral = Nothing
        Set GrupVari = Nothing
        Set GrupPrec = Nothing
        Set GrupPedi = Nothing
        Set GrupAlba = Nothing
        Set GrupEsta = Nothing
        
End Sub



Private Sub CargaMenuPublicidad(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Publicidad")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
End Sub




Private Sub CargaMenuCuotas(IdMenu As Integer)
        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Cuotas")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close

End Sub



Private Sub CargaMenuReparaciones(IdMenu As Integer)
        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "Reparaciones")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close

End Sub


Private Sub CargaMenuCRM(IdMenu As Integer)

        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "CRM")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
           
           
                    
                Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
        

End Sub


Private Sub CargaMenuUtilidades(IdMenu As Integer)
        'Creamos la TAB
        Set TabNuevo = RibbonBar.InsertTab(CLng(IdMenu), "UTILIDADES")
        TabNuevo.Id = CLng(IdMenu)
        Set GroupNew = TabNuevo.Groups.AddGroup("", 1000000)
        
        
        'todos los hijos que cuelgan en la tab
        Cad = "Select * from menus where aplicacion = 'aritaxi' and padre =" & IdMenu & " ORDER BY padre,orden"
        Rn2.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Cad = ""
        While Not Rn2.EOF
         
           If Not BloqueaPuntoMenu(Rn2!Codigo, "aritaxi") Then
                Habilitado = True
    
                If Not MenuVisibleUsuario(DBLet(Rn2!Codigo), "aritaxi") Then
                    Habilitado = False
                Else
                    If Not MenuVisibleUsuario(DBLet(Rn2!Padre), "aritaxi") Then Habilitado = False
                End If
           
           
                If DBLet(Rn2!Codigo, "N") = 1009 Then
                    '[Monica]01/02/2019: para el caso de sevilla no se llama facturacion electronica como en el resto
                    If vParamAplic.Cooperativa = 3 Then
                        Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, "Exportar Datos/Facturas")
                    Else
                        Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                    End If
                Else
                    Set Control = GroupNew.Add(xtpControlButton, Rn2!Codigo, Rn2!Descripcion)
                End If
                Control.Enabled = Habilitado
             
            End If
            Rn2.MoveNext
        Wend
        Rn2.Close
End Sub






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
        Case 101 ' Empresa
            frmConfParamGral.Show vbModal
        Case 102 ' Parámetros Aplicación
            Screen.MousePointer = vbHourglass
            Load frmConfParamAplic
            frmConfParamAplic.Show vbModal
        Case 103 ' Tipos de Movimiento
            frmConfTipoMov.Show vbModal
        Case 104 ' Tipos de Documentos
            frmConfParamRpt.Show vbModal
        Case 105 ' Usuarios
            frmMantenusu.Show vbModal
        Case 106 ' Seleccionar Impresora
            Screen.MousePointer = vbHourglass
            Me.CommonDialog1.ShowPrinter
            Screen.MousePointer = vbDefault
        Case 201 ' Marcas
            frmAlmMarcas.Show vbModal
        Case 202 ' Almacenes Propios
            frmAlmAlPropios.Show vbModal
        Case 203 ' Tipos Unidad
            frmAlmTipoUnidad.Show vbModal
        Case 204 ' Tipos artículos
            frmAlmTipoArticulo.Show vbModal
        Case 205 ' Familias artículos
            frmAlmFamiliaArticulo.Show vbModal
        Case 206 ' artículos
            frmAlmArticulos.Show vbModal
        Case 207 ' Movimientos Almacen
            frmAlmMovimientos.EsHistorico = False
            frmAlmMovimientos.hcoCodMovim = -1 'No carga el form al abrir
            frmAlmMovimientos.Show vbModal
        Case 208 ' Histórico Movimientos Almacen
            frmAlmMovimientos.EsHistorico = True
            frmAlmMovimientos.hcoCodMovim = -1
            frmAlmMovimientos.Show vbModal
        Case 209 ' Movimientos artículos
            frmAlmMovimArticulos.Show vbModal
        Case 211 ' Listado Movimientos
            AbrirListado (9)
        Case 212 ' Listado Artículos Inactivos
            AbrirListado (15)
        Case 213 ' Listado Artículos Componentes
            AbrirListado (11)
        Case 214 ' Listado Valoración Stocks
            AbrirListado (17)
        Case 215 ' Inf.stocks Máximos - Mínimos
            AbrirListado (18)
        Case 216 ' Inf.Stocks a una Fecha
            AbrirListado (19)
        Case 217 ' Toma inventario
            AbrirListado (12)
        Case 218 ' Entrada existencia real
            frmAlmInventario.Show vbModal
        Case 219 ' Listado diferencias
            AbrirListado (13)
        Case 220 ' Actualizar diferencias
            AbrirListado (14)
        Case 221 ' Valoración stocks inventariados
            AbrirListado (16)
        Case 222 ' Histórico inventario
            frmAlmHcoInven.Show vbModal
        Case 301 ' Clientes antes Actividades
'            frmFacActividades.Show vbModal
            frmFacClientes.Show vbModal
        Case 302 ' Textos Clientes Agrupados
            frmFacFormasEnvio.Show vbModal
        Case 303 ' Formas de Pago
            frmFacFormasPago.Show vbModal
        Case 304 ' Bancos Propios
            frmFacBancosPropios.Show vbModal
        Case 305 ' Situaciones especiales
            frmFacSituaciones.Show vbModal
        Case 306 ' Agentes comerciales
            frmFacAgentesCom.Show vbModal
        Case 307 ' Actividades , antes Clientes
'            frmFacClientes.Show vbModal
            frmFacActividades.Show vbModal
        Case 308 ' Tipos de Cartas
            frmFacCartasOferta.Show vbModal
        Case 309 ' Incidencias
            frmIncidencias.Show vbModal
        Case 310 ' Tarjetas
            frmTarjetas.Show vbModal
        Case 311 ' Clientes Inactivos
            AbrirListadoOfer (46)
        Case 312 ' Clientes
            AbrirListadoOfer (47)
        Case 313 ' Altas Clientes
            AbrirListadoOfer (48)
        Case 314 ' Etiquetas de clientes
            AbrirListadoOfer (90)
        Case 315 ' Cartas a clientes
            AbrirListadoOfer (91) '91: Informe Cartas a Clientes
        Case 316 ' Traspaso Taxitronic
            frmGesTraspaso.Show vbModal
        Case 317 ' Histórico de llamadas
            Select Case vParamAplic.Cooperativa
                Case 0, 2, 3
                    frmGesHisLlam.Show vbModal
                Case 1
                    frmGesHisLlamVIP.Show vbModal
            End Select
        Case 318 ' Mantenimiento Servicios Abonados
            frmGesServAbonados.Show vbModal
        Case 319 ' Facturación a Clientes
            frmFCliFacturac.Show vbModal
        Case 320 ' Facturas Varias a Clientes
            frmFCliFactuVar.Show vbModal
        Case 321 ' Histórico de Facturas
            frmFCliHcoFac.Show vbModal
        Case 322 ' Reimprimir Facturas
            frmFCliReImp.Show vbModal
        Case 323 ' Contabilizar Facturas
            frmFCliContaFac.Show vbModal
        Case 324 ' Facturas Rectificativas
            frmFCliRectif.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFCliRectif.hcoCodTipoM = "ARN" ' albaran rectificativo de cliente
            frmFCliRectif.EsHistorico = False
            frmFCliRectif.RecuperarFactu = False
            frmFCliRectif.Show vbModal
        Case 325 ' Ventas por Clientes
            AbrirListadoPed (230)
            BorrarTempInformes
        Case 326 ' Detalle Facturación
             AbrirListadoOfer (232)
        '[Monica]01/03/2018: informe de servicios de socios
        Case 401 ' Trabajadores
            frmAdmTrabajadores.Show vbModal
        Case 402 ' Vehículos
            frmGesVehic.Show vbModal
        Case 403 ' Chóferes
            frmGesConduc.Show vbModal
        Case 404 ' Socios
            frmGesSocios.Show vbModal
        Case 405 ' Histórico de Uves
            frmGesHcoUves.Show vbModal
        Case 406 ' Etiquetas de Socios
            AbrirListadoOfer 190
        Case 407 ' Certificados
            AbrirListadoOfer 242
        Case 408 ' Cartas a Socios
            AbrirListadoOfer 191
        Case 409 ' Mantenimiento Albaranes
            frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes.hcoCodTipoM = "ALV"
            frmFacEntAlbaranes.EsHistorico = False
            frmFacEntAlbaranes.RecuperarFactu = False
            frmFacEntAlbaranes.Show vbModal
        Case 410 ' Informe Albaranes por Artículo
            AbrirListadoPed (49)
        Case 411 ' Histórico Albaranes Anulados
        'Histórico de Albaranes eliminados
            frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes.hcoCodTipoM = "ALV"
            frmFacEntAlbaranes.EsHistorico = True
            frmFacEntAlbaranes.RecuperarFactu = False
            frmFacEntAlbaranes.Show vbModal
        Case 412 ' Previsión Facturación
            frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
            AbrirListadoPed (50) 'NO IMPRIME LISTADO
        Case 413 ' Facturación de Albaranes
            frmListadoPed.CodClien = "ALV" 'utilizamos esta vble para pasarle el tipo de movimiento
            AbrirListadoPed (52)
        Case 414 ' Facturas Rectificativas
            'Abre el formulario de Albaranes para introducir el Albaran Rectificativo
            'y desde este generar la Factura Rectificativa
            frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes.hcoCodTipoM = "ART"
            frmFacEntAlbaranes.EsHistorico = False
            frmFacEntAlbaranes.RecuperarFactu = False
            frmFacEntAlbaranes.Show vbModal
        Case 415 ' Histórico Albaran / Factura
            frmFacHcoFacturas2.hcoCodMovim = ""
            frmFacHcoFacturas2.publicidad = False
            frmFacHcoFacturas2.Show vbModal
        Case 416 ' Reimprimir Facturas
            AbrirListadoOfer 226
        Case 417 ' Contabilizar Facturas
            AbrirListado (223) 'Para pedir datos
        Case 418 ' Mantenimiento Servicios Socios
            frmGesServSocios.Show vbModal
        Case 419 ' Liquidación
            frmLiqLiquidaSoc.Show vbModal
        Case 420 ' Histórico Facturas
            frmLiqHcoFacSoc.Show vbModal
        Case 421 ' Reimprimir Facturas
            frmLiqReImp.Show vbModal
        Case 422 ' Contabilizar Facturas
            frmLiqContaFac.Show vbModal
        Case 423 ' Retenciones Socio
            frmLiqRetencion.Show vbModal
        Case 424 ' Ventas por Socio
            AbrirListadoPed (227)
            BorrarTempInformes
        Case 425 ' Ventas por meses
            AbrirListadoPed (229)
        Case 426 ' Ventas por Familia / Artículo
            AbrirListadoOfer (230)
        Case 427 ' Detalle Facturación
            AbrirListadoOfer (231)
        Case 428 ' Informe de servicios por socios
             AbrirListadoOfer (92)
        Case 429 ' Listado de liquidaciones Socios
             AbrirListadoOfer (241)
        Case 501 ' Proveedores
            frmComProveedores.Show vbModal
        Case 502 ' Proveedores Varios
            frmComProveV.Show vbModal
        Case 503 ' Direcciones
            frmComDirecciones.Show vbModal
        Case 504 ' Listado Proveedores
            AbrirListado (58)
        Case 505 ' Etiquetas de proveedores
            AbrirListadoOfer (305)
        Case 506 ' Cartas a Proveedores
            AbrirListadoOfer (306) '306: Informe Cartas a Proveedores
        Case 507 ' Precios Proveedor
            frmComPreciosProv.Show vbModal
        Case 508 ' Descuentos Proveedor
            frmComDtosFamMarca.Show vbModal
        Case 509 ' Mto.Pedidos Proveedor
            frmComEntPedidos.MostrarDatos = ""
            frmComEntPedidos.EsHistorico = False
            frmComEntPedidos.Show vbModal
        Case 510 ' Histórico Pedidos Anulados
            frmComEntPedidos.MostrarDatos = ""
            frmComEntPedidos.EsHistorico = True
            frmComEntPedidos.Show vbModal
        Case 511 ' List.Material pendiente de recibir
            AbrirListadoOfer (307) '307: List. Materia pte recibir
        Case 512 ' Mant.Albaranes Proveedor
        'Mantenimiento de Albaranes a Proveedor
            frmComEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmComEntAlbaranes.EsHistorico = False
            frmComEntAlbaranes.Show vbModal
        Case 513 ' Histórico Albaranes Anulados
            frmComEntAlbaranes.EsHistorico = True
            frmComEntAlbaranes.Show vbModal
        Case 514 ' Inf.Pendiente de facturar
            AbrirListadoOfer (308)
        Case 515 ' Recepción Facturas
            frmComFacturar.Show vbModal
        Case 516 ' Histórico Albaran / Factura
            frmComHcoFacturas.hcoCodMovim = ""
            frmComHcoFacturas.Show vbModal
        Case 517 ' Contabilizar Facturas
            AbrirListado (224) 'Para pedir datos
        Case 518 ' Compras por Proveedor
            AbrirListadoOfer (310)
        Case 519 ' Compras por Familia/Artic.
            AbrirListadoOfer (311)
        Case 520 ' Albaranes por Proveedor
            AbrirListadoOfer (312)
        Case 601 ' Facturación Clientes
            frmPubliFacCli.Show vbModal
        Case 602 ' Facturas Rectificativas
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
        Case 603 ' Histórico Facturas Clientes
            frmPubliHcoFacCli.Show vbModal
        Case 604 ' Facturación Socios
            frmPubliFacSoc.Show vbModal
        Case 605 ' Facturas Rectificativas Socios
            frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes.hcoCodTipoM = "ARQ"
            frmFacEntAlbaranes.EsHistorico = False
            frmFacEntAlbaranes.RecuperarFactu = False
            frmFacEntAlbaranes.Show vbModal
        Case 606 ' Histórico Facturas Socios
            frmPubliHcoFacSoc.Show vbModal
        Case 607 ' Reimprimir Facturas
            frmPubliReImp.Show vbModal
        Case 608 ' Contabilizar Facturas
            frmPubliContaFac.Show vbModal
        Case 701 ' Generar Facturas Cuotas
            frmCuotasFac.Show vbModal
        Case 702 ' Reimprimir Facturas
            frmCuotasReImp.Show vbModal
        Case 703 ' Histórico Facturas
            frmCuotasHcoFacturas.Show vbModal
        Case 704 ' Contabilizar Facturas
            frmCuotasContaFac.Show vbModal
        Case 705 ' Mantenimiento Albaranes
            frmFacEntAlbaranes.hcoCodTipoM = "ALS"
            frmFacEntAlbaranes.Show vbModal
        Case 706 ' Previsión Facturación
            frmListadoPed.CodClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
            AbrirListadoPed (50) 'NO IMPRIME LISTADO
        Case 707 ' Facturación
            frmListadoPed.CodClien = "ALS" 'utilizamos esta vble para pasarle el tipo de movimiento
            AbrirListadoPed (52)
        Case 708 ' Facturas Rectificativas
            frmFacEntAlbaranes.hcoCodMovim = "" 'No carga el form con datos al abrir
            frmFacEntAlbaranes.hcoCodTipoM = "ARC"
            frmFacEntAlbaranes.EsHistorico = False
            frmFacEntAlbaranes.RecuperarFactu = False
            frmFacEntAlbaranes.Show vbModal
        
        Case 801 ' Mant.Reparaciones
            'Mantenimiento de Reparaciones
            frmRepEntReparaciones.EntradaEquipo = ""
            frmRepEntReparaciones.ControlRep = False
            frmRepEntReparaciones.EsHistorico = False
            frmRepEntReparaciones.Show vbModal
        
        Case 802 ' Control Reparaciones
            'Control de Reparaciones (para los Tecnicos)
            frmRepEntReparaciones.EntradaEquipo = ""
            frmRepEntReparaciones.ControlRep = True
            frmRepEntReparaciones.EsHistorico = False
            frmRepEntReparaciones.Show vbModal
        
        Case 803 ' Mant.Nº Serie
           'Mantenimiento de Nºs de Serie
            frmRepNumSerie2.Show vbModal
        
        Case 804 ' Motivos baja equipos
           'Motivos baja equipos
            frmRepMotivosBaja.Show vbModal
        
        Case 805 ' Motivos Pend.Reparación
           'Motivos Pendientes Reparar
            frmRepMotivosPend.Show vbModal
                    
        Case 806 ' Servicios asistencia técnica
            frmManSat.Show vbModal
        Case 807 ' Tipos averia
            frmtipave.Show vbModal
        Case 809 ' Trabajos realizados
            frmManTraReali.Show vbModal
        Case 810 ' Listado Reparaciones del Dia
            AbrirListado2 1
        
        
        Case 901 ' Mantenimiento acciones comerciales
            frmCRMMto.DesdeElCliente = 0 'No clien
            frmCRMMto.TipoPredefinido = 0   'Ninguno
            frmCRMMto.Show vbModal
        Case 902 ' Tipos acciones comerciales
            frmCRMtipos.Show vbModal
        Case 903 ' Generar acciones comerciales
            frmCRMVarios.Opcion = 0
            frmCRMVarios.Show vbModal
        
        Case 1001 'exportar email a csv
            AbrirListado 101
' Avisos
'            If TieneAvisosPendientes Then
'                frmAlertas.Show vbModal
'            Else
'                MsgBox "No hay avisos para mostrar", vbInformation
'            End If
        
        Case 1002
            frmGesHisLlamAnt.Show vbModal
' llamadas
'            frmLlamadas.Show vbModal
        Case 1003 ' Concepto llamadas
            frmLlamadasTipo.Show vbModal
        Case 1004 ' Copia Seguridad local
            frmBackUP.Show vbModal
        Case 1005 ' Borre Facturas y Movimiento
            AbrirListado 97
        Case 1006 ' Revisar caracteres especiales
            AbrirListado2 3
        Case 1007 ' Acciones realizadas
            Screen.MousePointer = vbHourglass
            Load frmLog
            DoEvents
            frmLog.Show vbModal
            Screen.MousePointer = vbDefault
        Case 1008 ' Eliminar artículos
            frmVarios.Opcion = 1
            frmVarios.Show vbModal
        Case 1009 ' Facturacion Electrónica
            If vParamAplic.Cooperativa = 3 Then
                frmExportarFacturasSev.Show vbModal
            Else
                frmExportarFacturas.Show vbModal
            End If
        Case 1010 ' Errores en NºFactura Cliente
            'Buscar errores en nº de factura (solo en facturas de clientes)
            Screen.MousePointer = vbHourglass
            frmUtilidades.Opcion = 5
            frmUtilidades.Show vbModal

'            frmPrueba.Show vbModal
            

        Case 1011 ' Fras Pdtes Contabilizar Clientes
            'Facturas pendientes de contabilizar (CLIENTES)
            Screen.MousePointer = vbHourglass
            frmUtilidades.Opcion = 6
            frmUtilidades.Show vbModal
        
        Case 1012 ' Fras Pdtes Contabilizar Proveedores
            'Facturas pendientes de contabilizar (PROVEEDORES)
            Screen.MousePointer = vbHourglass
            frmUtilidades.Opcion = 7
            frmUtilidades.Show vbModal
        
        Case 1013 ' Usuarios activos
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
                    
        
        
        Case 1014 ' Conexiones activas
            'ver las conexiones a donde apuntan
            Dim Cad As String
             
                
                MostrarCadenasConexion
                
                
        Case Else
  
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


Private Sub mnHerrAriadnaCC_Click(Index As Integer)
 
'        If vUsu.Nivel > 1 Then
'            MsgBox "No tiene permisos", vbExclamation
'            Exit Sub
'        End If
'        'El index 3 , que es la barra, en frmCC es la opcion de NUEVA EMPRESA
'        ' y no se llma desde aqui, con lo cual no hay problemo
'        'Para el restro cojo el valor del helpidi
'
'        frmCentroControl.Opcion = Index
'        frmCentroControl.Show vbModal
    
End Sub



'Establecer y fijar Skin
Public Sub EstablecerSkin(QueSkin As Integer)

    FijaSkin QueSkin

  ' Cargando el archivo del Skin
  ' ============================
    'frmPpal.SkinFramework1.LoadSkin Skn$, ""
    Me.SkinFramework1.ApplyWindow frmppal.hwnd
    Me.SkinFramework1.ApplyOptions = Me.SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics
    


    
End Sub

Private Function FijaSkin(numero)
    Me.SkinFramework1.ExcludeModule "crviewer9.dll"

  Select Case (numero)
 
           
            Case 1:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlue.ini"
            Case 2:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalSilver.ini"
            Case 3:
                Skn$ = CStr(App.Path & "\Styles\Office2010.cjstyles")
                Me.SkinFramework1.LoadSkin Skn$, "NormalBlack.ini"
                
                  
                
        
        
  End Select
    
End Function



Private Sub PonerCaption()
        Caption = "AriTAXI 6    V-" & App.Major & "." & App.Minor & "." & App.Revision & "    usuario: " & vUsu.NOMBRE
        'Label33.Caption = "   " & vEmpresa.nomempre
End Sub


Public Sub OpcionesMenuInformacion(Id As Long)
    
    Select Case Id
    Case ID_Licencia_Usuario_Final_txt
        LanzaVisorMimeDocumento Me.hwnd, "c:\programas\Ariadna.rtf"
    Case ID_Licencia_Usuario_Final_web
        LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "AriTAXI-6.html?Licenciadeuso.html"
    Case ID_Ver_Version_operativa_web
        LanzaVisorMimeDocumento Me.hwnd, DireccionAyuda & "Aritaxi-6.html"  ' "http://www.ariadnasw.com/clientes/"
    End Select
    
End Sub

Public Function BloqueaPuntoMenu(IdProg As Long, Aplicacion As String) As Boolean
Dim EsdeAnalitica As Boolean

    BloqueaPuntoMenu = False

    If Aplicacion = "aritaxi" Then
        ' programas de analitica
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

Private Sub AbrirListado2(KOpcion As Integer)
    Screen.MousePointer = vbHourglass
    frmListado2.Opcion = KOpcion
    frmListado2.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

