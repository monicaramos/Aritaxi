VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdmTrabajadores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trabajadores"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "frmAdmTrabajadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   240
      TabIndex        =   80
      Top             =   410
      Width           =   11175
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   885
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "C�digo Trabajador|N|N|0|9999|straba|codtraba|0000|S|"
         Text            =   "Text"
         Top             =   200
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2800
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Nombre Trabajador|T|N|||straba|nomtraba||N|"
         Text            =   "Text1"
         Top             =   200
         Width           =   4485
      End
      Begin VB.Label Label1 
         Caption         =   "C�digo"
         Height          =   255
         Index           =   0
         Left            =   280
         TabIndex        =   82
         Top             =   200
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   81
         Top             =   200
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   9960
      TabIndex        =   28
      Top             =   5520
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   5360
      Width           =   2895
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   180
         Width           =   2715
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9930
      TabIndex        =   29
      Top             =   5520
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   8640
      TabIndex        =   27
      Top             =   5520
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3120
      Top             =   5400
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Estudios/Formaci�n"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Habilidades"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Experiencia Laboral"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Formaci�n Realizada"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Formaci�n Empresa"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   8400
         TabIndex        =   33
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   4560
      Top             =   5400
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
      Height          =   4215
      Left            =   240
      TabIndex        =   34
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Datos b�sicos"
      TabPicture(0)   =   "frmAdmTrabajadores.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(34)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(15)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(36)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(37)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "imgBuscar(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(12)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(24)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(25)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(26)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "imgBuscar(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ImgMail(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "imgBuscar(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text1(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text1(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text1(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "frameBancos"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "frameDptoPersonal"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text1(8)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text1(23)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text1(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text1(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text1(24)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text2(24)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text2(10)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(10)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Estudios/Formaci�n"
      TabPicture(1)   =   "frmAdmTrabajadores.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(1)=   "txtAux1(0)"
      Tab(1).Control(2)=   "txtAux1(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Habilidades"
      TabPicture(2)   =   "frmAdmTrabajadores.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtAux2"
      Tab(2).Control(1)=   "DataGrid2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Experiencia Laboral"
      TabPicture(3)   =   "frmAdmTrabajadores.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtAux3(1)"
      Tab(3).Control(1)=   "TxtAux3(0)"
      Tab(3).Control(2)=   "DataGrid3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Formaci�n Realizada"
      TabPicture(4)   =   "frmAdmTrabajadores.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TxtAux4(4)"
      Tab(4).Control(1)=   "TxtAux4(3)"
      Tab(4).Control(2)=   "TxtAux4(2)"
      Tab(4).Control(3)=   "TxtAux4(1)"
      Tab(4).Control(4)=   "TxtAux4(0)"
      Tab(4).Control(5)=   "DataGrid4"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Formaci�n Empresa"
      TabPicture(5)   =   "frmAdmTrabajadores.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "DataGrid5"
      Tab(5).Control(1)=   "TxtAux5(0)"
      Tab(5).Control(2)=   "TxtAux5(1)"
      Tab(5).Control(3)=   "TxtAux5(2)"
      Tab(5).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   10
         Left            =   1365
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "Centro de coste|T|S|||straba|codccost||N|"
         Top             =   3750
         Width           =   630
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   10
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   83
         Text            =   "Text2"
         Top             =   3750
         Width           =   3340
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Height          =   315
         Index           =   24
         Left            =   2030
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   79
         Text            =   "Text2"
         Top             =   3351
         Width           =   3340
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   24
         Left            =   1365
         MaxLength       =   3
         TabIndex        =   11
         Tag             =   "Almacen por Defecto|N|N|0|999|straba|codalmac|000|N|"
         Text            =   "Text aldu dkdo s�s�s"
         Top             =   3351
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Domicilio|T|N|||straba|domtraba||N|"
         Text            =   "Text1"
         Top             =   988
         Width           =   4365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   7
         Left            =   1020
         MaxLength       =   15
         TabIndex        =   8
         Tag             =   "Tel�fono|T|N|||straba|teltraba||N|"
         Text            =   "Text1"
         Top             =   2152
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   23
         Left            =   3645
         MaxLength       =   20
         TabIndex        =   3
         Tag             =   "Login Trabajador|T|S|||straba|login||N|"
         Text            =   "Text aldu dkdo s�s�s"
         Top             =   600
         Width           =   1710
      End
      Begin VB.TextBox TxtAux5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -68160
         MaxLength       =   15
         TabIndex        =   76
         Tag             =   "Resultado|T|N|||strab5|resforma||N|"
         Text            =   "Resultado"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox TxtAux5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -72840
         MaxLength       =   50
         TabIndex        =   75
         Tag             =   "Formaci�n|T|N|||strab5|formaci2||N|"
         Text            =   "Formacion"
         Top             =   3720
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.TextBox TxtAux5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74400
         MaxLength       =   12
         TabIndex        =   74
         Tag             =   "Fecha Formaci�n|F|N|||strab5|fecforma||N|"
         Text            =   "F. Formac."
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtAux4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   4
         Left            =   -67200
         MaxLength       =   8
         TabIndex        =   73
         Tag             =   "Evaluaci�n|T|N|||strab4|evaluaci||N|"
         Text            =   "Evaluaci"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox TxtAux4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   -69000
         MaxLength       =   15
         TabIndex        =   72
         Tag             =   "Centro|T|N|||strab4|codcentr||N|"
         Text            =   "Centro"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox TxtAux4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   -71760
         MaxLength       =   40
         TabIndex        =   71
         Tag             =   "Formacion|T|N|||strab4|formaci1||N|"
         Text            =   "Formacion"
         Top             =   3720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox TxtAux4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -72960
         MaxLength       =   12
         TabIndex        =   70
         Tag             =   "Fecha Evaluaci�n|F|N|||strab4|fecevalu|dd/mm/yyyy|N|"
         Text            =   "F.Evalua."
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtAux4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74280
         MaxLength       =   12
         TabIndex        =   69
         Tag             =   "Fecha Formaci�n|F|N|||strab4|fecforma|dd/mm/yyyy|N|"
         Text            =   "F.Formac."
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TxtAux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -72360
         MaxLength       =   70
         TabIndex        =   68
         Tag             =   "Experiencia|T|N|||strab3|experien||N|"
         Text            =   "Experiencia"
         Top             =   3720
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.TextBox TxtAux3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74400
         MaxLength       =   15
         TabIndex        =   67
         Tag             =   "Periodo|T|N|||strab3|periodo1||N|"
         Text            =   "Periodo"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -74160
         MaxLength       =   70
         TabIndex        =   66
         Tag             =   "Habilidad|T|N|||strab2|habilida||N|"
         Text            =   "Habilidad"
         Top             =   3720
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   9
         Tag             =   "Cargo en la empresa|T|S|||straba|cartraba||N|"
         Text            =   "Text1"
         Top             =   2540
         Width           =   4365
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   -72120
         MaxLength       =   70
         TabIndex        =   60
         Tag             =   "Formaci�n|T|N|||strab1|formacio||N|"
         Text            =   "Formacion Formacion Formacion Formacion Formacion Formacion Formacion "
         Top             =   3660
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   0
         Left            =   -74160
         MaxLength       =   15
         TabIndex        =   59
         Tag             =   "Periodo|T|N|||strab1|periodos||N|"
         Text            =   "Periodo"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame frameDptoPersonal 
         Caption         =   "Datos relacionados con Dpto Personal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   5760
         TabIndex        =   41
         Top             =   540
         Width           =   5175
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   13
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   16
            Tag             =   "Fecha de Baja|F|S|||straba|fechabaj|dd/mm/yyyy|N|"
            Top             =   840
            Width           =   1040
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   12
            Left            =   3960
            MaxLength       =   10
            TabIndex        =   15
            Tag             =   "Fecha de Alta|F|N|||straba|fechaalt|dd/mm/yyyy|N|"
            Top             =   420
            Width           =   1040
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   11
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   13
            Tag             =   "Fecha de Nacimiento|F|N|||straba|fechanac|dd/mm/yyyy|N|"
            Top             =   420
            Width           =   1040
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   14
            Left            =   1040
            MaxLength       =   12
            TabIndex        =   14
            Tag             =   "N� SSocial|T|S|||straba|nrosegur||N|"
            Text            =   "Text1"
            Top             =   840
            Width           =   1440
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Baja"
            Height          =   255
            Index           =   5
            Left            =   2760
            TabIndex        =   48
            Top             =   840
            Width           =   855
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   2
            Left            =   3675
            Picture         =   "frmAdmTrabajadores.frx":00B4
            ToolTipText     =   "Buscar fecha"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Alta"
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   47
            Top             =   420
            Width           =   855
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   3675
            Picture         =   "frmAdmTrabajadores.frx":013F
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "F. Nacimiento"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   46
            Top             =   420
            Width           =   975
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   0
            Left            =   1155
            Picture         =   "frmAdmTrabajadores.frx":01CA
            ToolTipText     =   "Buscar fecha"
            Top             =   450
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "N� S. Social"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   975
         End
      End
      Begin VB.Frame frameBancos 
         Caption         =   "Datos Bancarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   1635
         Left            =   5790
         TabIndex        =   43
         Top             =   2400
         Width           =   5175
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   26
            Left            =   960
            MaxLength       =   4
            TabIndex        =   22
            Tag             =   "IBAN Gastos1|T|S|||straba|iban1||N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   25
            Left            =   960
            MaxLength       =   4
            TabIndex        =   17
            Tag             =   "IBAN|T|S|||straba|iban||N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   15
            Left            =   1650
            MaxLength       =   4
            TabIndex        =   18
            Tag             =   "C�digo Banco N�mina|N|S|0|9999|straba|codbanco|0000|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   16
            Left            =   2325
            MaxLength       =   4
            TabIndex        =   19
            Tag             =   "Sucursal N�mina|N|S|0|9999|straba|codsucur|0000|N|"
            Text            =   "Text1"
            Top             =   480
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   17
            Left            =   3015
            MaxLength       =   2
            TabIndex        =   20
            Tag             =   "D�gito Control N�mina|T|S|||straba|digcontr|00||"
            Text            =   "Text1"
            Top             =   480
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   18
            Left            =   3450
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "Cuenta Bancaria N�mina|T|S|||straba|cuentaba|0000000000||"
            Text            =   "Text1"
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   19
            Left            =   1650
            MaxLength       =   4
            TabIndex        =   23
            Tag             =   "C�digo Banco Gastos|N|N|0|9999|straba|codbanc1|0000|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   20
            Left            =   2325
            MaxLength       =   4
            TabIndex        =   24
            Tag             =   "Sucursal Gastos|N|S|0|9999|straba|codsucu1|0000|N|"
            Text            =   "Text1"
            Top             =   1200
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   21
            Left            =   3015
            MaxLength       =   2
            TabIndex        =   25
            Tag             =   "D�gito Control Gastos|T|S|||straba|digcont1|00||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   405
         End
         Begin VB.TextBox Text1 
            Height          =   315
            Index           =   22
            Left            =   3450
            MaxLength       =   10
            TabIndex        =   26
            Tag             =   "Cuenta Bancaria Gastos|T|S|||straba|cuentab1|0000000000||"
            Text            =   "Text1"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   8
            Left            =   960
            TabIndex        =   86
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "IBAN"
            Height          =   255
            Index           =   4
            Left            =   960
            TabIndex        =   85
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
            Height          =   255
            Index           =   39
            Left            =   1650
            TabIndex        =   57
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   35
            Left            =   2325
            TabIndex        =   56
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
            Height          =   255
            Index           =   33
            Left            =   3090
            TabIndex        =   55
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Bancaria"
            Height          =   255
            Index           =   29
            Left            =   3450
            TabIndex        =   54
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "N�mina"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   53
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
            Height          =   255
            Index           =   19
            Left            =   1650
            TabIndex        =   52
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Index           =   17
            Left            =   2325
            TabIndex        =   51
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "DC"
            Height          =   255
            Index           =   7
            Left            =   3090
            TabIndex        =   50
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Bancaria"
            Height          =   255
            Index           =   6
            Left            =   3450
            TabIndex        =   49
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Gastos"
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   43
            Left            =   240
            TabIndex        =   44
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   9
         Left            =   1005
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "e-mail|T|S|||straba|maitraba||N|"
         Text            =   "Text1"
         Top             =   2930
         Width           =   4365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   6
         Left            =   1005
         MaxLength       =   9
         TabIndex        =   2
         Tag             =   "N.I.F.|T|N|||straba|niftraba||N|"
         Text            =   "Text1"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   5
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "Provincia|T|N|||straba|protraba||N|"
         Text            =   "Text1"
         Top             =   1764
         Width           =   4365
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   2910
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Poblaci�n|T|N|||straba|pobtraba||N|"
         Text            =   "Text1"
         Top             =   1376
         Width           =   2460
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "C.Postal|T|N|||straba|codpobla||N|"
         Text            =   "Text1"
         Top             =   1376
         Width           =   825
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmAdmTrabajadores.frx":0255
         Height          =   3510
         Left            =   -74520
         TabIndex        =   58
         Top             =   520
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmAdmTrabajadores.frx":026A
         Height          =   3510
         Left            =   -74520
         TabIndex        =   62
         Top             =   520
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid3 
         Bindings        =   "frmAdmTrabajadores.frx":027F
         Height          =   3510
         Left            =   -74520
         TabIndex        =   63
         Top             =   520
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "frmAdmTrabajadores.frx":0294
         Height          =   3510
         Left            =   -74520
         TabIndex        =   64
         Top             =   520
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid5 
         Bindings        =   "frmAdmTrabajadores.frx":02A9
         Height          =   3510
         Left            =   -74520
         TabIndex        =   65
         Top             =   520
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6191
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         BorderStyle     =   0
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "CCoste"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   84
         Top             =   3750
         Width           =   615
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1005
         Picture         =   "frmAdmTrabajadores.frx":02BE
         Tag             =   "-1"
         ToolTipText     =   "Buscar centro coste"
         Top             =   3750
         Width           =   240
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   0
         Left            =   705
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   2930
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1005
         Picture         =   "frmAdmTrabajadores.frx":03C0
         Tag             =   "-1"
         ToolTipText     =   "Buscar almacen"
         Top             =   3351
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Index           =   26
         Left            =   240
         TabIndex        =   78
         Top             =   3351
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
         Height          =   255
         Index           =   25
         Left            =   2880
         TabIndex        =   77
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cargo"
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   61
         Top             =   2540
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Tel�fono"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   45
         Top             =   2152
         Width           =   735
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1005
         Picture         =   "frmAdmTrabajadores.frx":04C2
         Tag             =   "-1"
         ToolTipText     =   "Buscar poblaci�n"
         Top             =   1376
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
         Height          =   255
         Index           =   37
         Left            =   240
         TabIndex        =   40
         Top             =   2930
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "N.I.F."
         Height          =   255
         Index           =   36
         Left            =   240
         TabIndex        =   39
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   15
         Left            =   240
         TabIndex        =   38
         Top             =   1764
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Poblaci�n"
         Height          =   255
         Index           =   34
         Left            =   2160
         TabIndex        =   37
         Top             =   1376
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "C. Postal"
         Height          =   255
         Index           =   14
         Left            =   240
         TabIndex        =   36
         Top             =   1376
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   35
         Top             =   988
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   5880
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   7320
      Top             =   5400
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
   Begin MSAdodcLib.Adodc Data5 
      Height          =   330
      Left            =   3120
      Top             =   5640
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
   Begin MSAdodcLib.Adodc Data6 
      Height          =   330
      Left            =   4560
      Top             =   5640
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
   Begin VB.Menu mnMtoLineas 
      Caption         =   "&Mantenimiento Lineas"
      Begin VB.Menu mnEstudios 
         Caption         =   "&Estudios/Formaci�n"
         HelpContextID   =   2
      End
      Begin VB.Menu mnHabilidades 
         Caption         =   "&Habilidades"
         HelpContextID   =   2
      End
      Begin VB.Menu mnExperiencia 
         Caption         =   "Experiencia &Laboral"
         HelpContextID   =   2
      End
      Begin VB.Menu mnFormRealizada 
         Caption         =   "&Formaci�n Realizada"
         HelpContextID   =   2
      End
      Begin VB.Menu mnFormEmpresa 
         Caption         =   "F&ormacion Empresa"
         HelpContextID   =   2
      End
   End
End
Attribute VB_Name = "frmAdmTrabajadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmCP As frmCPostal 'Codigos Postales
Attribute frmCP.VB_VarHelpID = -1
Private WithEvents frmAlm As frmAlmAlPropios  'Almacenes Propios
Attribute frmAlm.VB_VarHelpID = -1

Private Modo As Byte
'-------------------------------------------------------
'Se distinguen varios MODOS
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------

Dim ModificaLineas As Byte
'1.- A�adir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim NumTabMto As Byte 'Indica que numero de Tab que esta en modo Mantenimiento

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la tabla principal del formulario
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas del Mantenimiento en que estemos

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnAnyadir As Byte
'Variable que indica el n�mero del Boton  Anyadir en la Toolbar1

'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1
Dim btnPrimero As Byte

Dim BuscaChekc As String

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos


'===========================================================================
'       PROCEDIMIENTOS
'============================================================================

Private Sub cmdAceptar_Click()
Dim cad As String, Indicador As String

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
                    TerminaBloquear
                    PosicionarData
                End If
            End If
                
         Case 5 'INSERTAR MODIFICAR LINEA
            'Actualizar el registro en la tabla de lineas 'sdirec' (Direcciones/Departamentos)
            cad = "Select * from " & NomTablaLineas & " where codtraba= " & Data1.Recordset!CodTraba
            cad = cad & " order by numlinea"
            
            If ModificaLineas = 1 Then 'INSERTAR lineas
                If InsertarLinea Then
                    Select Case Me.SSTab1.Tab
                        Case 1 'Estudios/Formacion - Datos de la tabla strab1
                            CargaGrid DataGrid1, Data2, cad
                        Case 2 'Habilidades
                            CargaGrid DataGrid2, Data3, cad
                        Case 3 'Experiencia Laboral
                            CargaGrid DataGrid3, Data4, cad
                        Case 4 'Formacion Realizada
                            CargaGrid DataGrid4, Data5, cad
                        Case 5 'Formacion Empresa
                            CargaGrid DataGrid5, Data6, cad
                    End Select
                    BotonAnyadirLinea
                End If
                
            ElseIf ModificaLineas = 2 Then 'MODIFICAR lineas
                If ModificarLinea Then
                    TerminaBloquear
                    PonerBotonCabecera True
                    ModificaLineas = 0
                    Select Case Me.SSTab1.Tab
                        Case 1 'Estudios/Formacion - Datos de la tabla strab1
                            NumRegElim = Data2.Recordset.AbsolutePosition
                            CargaTxtAux1 False, False
                            'CargaGrid DataGrid1, Data2, cad
                            CargaGrid2 DataGrid1, Data2
                            SituarDataPosicion Data2, NumRegElim, Indicador
                        Case 2 'Habilidades
                            NumRegElim = Data3.Recordset.AbsolutePosition
                            CargaTxtAux2 False, False
                            'CargaGrid DataGrid2, Data3, cad
                            CargaGrid2 DataGrid2, Data3
                            SituarDataPosicion Data3, NumRegElim, Indicador
                        Case 3 'Experiencia Laboral
                            NumRegElim = Data4.Recordset.AbsolutePosition
                            CargaTxtAux3 False, False
                            CargaGrid2 DataGrid3, Data4
                            SituarDataPosicion Data4, NumRegElim, Indicador
                        Case 4 'Formacion Realizada
                            NumRegElim = Data5.Recordset.AbsolutePosition
                            CargaTxtAux4 False, False
                            CargaGrid2 DataGrid4, Data5
                            SituarDataPosicion Data5, NumRegElim, Indicador
                        Case 5 'Formacion Empresa
                            NumRegElim = Data6.Recordset.AbsolutePosition
                            CargaTxtAux5 False, False
                            CargaGrid2 DataGrid5, Data6
                            SituarDataPosicion Data6, NumRegElim, Indicador
                    End Select
                End If
            End If
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()

    Select Case Modo
        Case 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            Select Case Me.SSTab1.Tab
                Case 1 'Estudios/Formacion
                    CargaTxtAux1 False, False
                    DataGrid1.Enabled = True
                    If ModificaLineas = 1 Then 'Insertar
                        DataGrid1.AllowAddNew = False
                        If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
                    End If
                Case 2 'Habilidades
                    CargaTxtAux2 False, False
                    DataGrid2.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid2.AllowAddNew = False
                        If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                    End If
                Case 3 'Experiencia Laboral
                    CargaTxtAux3 False, False
                    DataGrid3.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid3.AllowAddNew = False
                        If Not Data4.Recordset.EOF Then Data4.Recordset.MoveFirst
                    End If
                Case 4 'Formacion Realizada
                    CargaTxtAux4 False, False
                    DataGrid4.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid4.AllowAddNew = False
                        If Not Data5.Recordset.EOF Then Data5.Recordset.MoveFirst
                    End If
                Case 5 'Formacion Empresa
                    CargaTxtAux5 False, False
                    DataGrid5.Enabled = True
                    If ModificaLineas = 1 Then 'INSERTAR
                        DataGrid5.AllowAddNew = False
                        If Not Data6.Recordset.EOF Then Data6.Recordset.MoveFirst
                    End If
            End Select
            PonerBotonCabecera True
            ModificaLineas = 0
    End Select
End Sub


Private Sub BotonAnyadir()
'A�adir registro en tabla de trabajadores: straba (Cabecera)

    LimpiarCampos 'Vac�a los TextBox
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 3
    
    Text1(0).Text = SugerirCodigoSiguienteStr("straba", "codtraba")
    Text1(12).Text = Format(Now, "dd/mm/yyyy")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonAnyadirLinea()
    'Si no estaba modificando lineas salimos
    ' Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 2 Then Exit Sub
        
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede A�adir. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    ModificaLineas = 1 'Ponemos Modo A�adir Linea
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerBotonCabecera False
    lblIndicador.Caption = "INSERTAR"
    
    Select Case Me.SSTab1.Tab
        Case 1 'Estudios / Formacion
                'Situamos el grid al final
                AnyadirLinea DataGrid1, Data2
                CargaTxtAux1 True, True
                PonerFoco txtAux1(0)
        Case 2 'Habilidades
                AnyadirLinea DataGrid2, Data3
                CargaTxtAux2 True, True
                PonerFoco txtAux2
        Case 3 'Experiencia Laboral
                AnyadirLinea DataGrid3, Data4
                CargaTxtAux3 True, True
                PonerFoco txtAux3(0)
        Case 4 'Formacion Realizada
                AnyadirLinea DataGrid4, Data5
                CargaTxtAux4 True, True
                PonerFoco txtAux4(0)
        Case 5 'Formacion Empresa
                AnyadirLinea DataGrid5, Data6
                CargaTxtAux5 True, True
                PonerFoco TxtAux5(0)
    End Select
End Sub


Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
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


Private Sub BotonVerTodos()
'Ver todos
    LimpiarCampos
    Me.SSTab1.Tab = 0
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonModificar()
    'A�adiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4

    PonerFoco Text1(1)
End Sub


Private Sub BotonModificarLinea()
'Modificar una linea
Dim vWhere As String

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then Exit Sub '1= Insertar
    
    If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede Modificar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If
    
    vWhere = "codtraba=" & Val(Text1(0).Text) & " and numlinea="
    Select Case Me.SSTab1.Tab
        Case 1 'Estudios/Formacion
                If Data2.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data2.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux1 True, False
                DataGrid1.Enabled = False
                PonerFoco txtAux1(0)
        Case 2 'Habilidades
                If Data3.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data3.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux2 True, False
                DataGrid2.Enabled = False
                PonerFoco txtAux2
        Case 3 'Experiencia Laboral
                If Data4.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data4.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux3 True, False
                DataGrid3.Enabled = False
                PonerFoco txtAux3(0)
        Case 4 'Formacion Realizada
                If Data5.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data5.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux4 True, False
                DataGrid4.Enabled = False
                PonerFoco txtAux4(0)
        Case 5 'Formacion Empresa
                If Data6.Recordset.EOF Then Exit Sub
                vWhere = vWhere & Data6.Recordset!numlinea
                If Not BloqueaRegistro(NomTablaLineas, vWhere) Then Exit Sub
                CargaTxtAux5 True, False
                DataGrid5.Enabled = False
                PonerFoco TxtAux5(0)
    End Select
    
    ModificaLineas = 2 'Modificar
    'A�adiremos el boton de aceptar y demas objetos para insertar
    Me.lblIndicador.Caption = "MODIFICAR"
    PonerBotonCabecera False
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de trabajadores (straba)
Dim cad As String
On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
    cad = "Cabecera de Trabajadores." & vbCrLf
    cad = cad & "------------------------------" & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Trabajador:"
    cad = cad & vbCrLf & "C�digo:   " & Format(Data1.Recordset.Fields(0), "000000")
    cad = cad & vbCrLf & "Descripci�n:   " & Data1.Recordset.Fields(1)
    cad = cad & vbCrLf & vbCrLf & " �Desea Eliminarlo? "
    
    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        
        If Not Eliminar Then
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Trabajador", Err.Description
End Sub


Private Sub BotonEliminarLinea()
'Eliminar una linea De Trabajador. Tablas: strab1, strab2, strab3, strab4, strab5
Dim Sql As String
Dim numlinea As Integer
On Error GoTo EEliminarLinea

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Or ModificaLineas = 2 Then Exit Sub '1= Insertar, 2=Modificar

     If NumTabMto <> Me.SSTab1.Tab Then
        MsgBox "No puede eliminar. Esta en Modo Mantenimiento de otra linea.", vbExclamation
        Exit Sub
    End If

    Select Case Me.SSTab1.Tab
        Case 1 'EStudios/Formacion
            If Data2.Recordset.EOF Then Exit Sub
            numlinea = Data2.Recordset!numlinea
        Case 2 'Habilidades
            If Data3.Recordset.EOF Then Exit Sub
            numlinea = Data3.Recordset!numlinea
        Case 3 'Experiencia Laboral
            If Data4.Recordset.EOF Then Exit Sub
            numlinea = Data4.Recordset!numlinea
        Case 4 'Formacion Realizada
            If Data5.Recordset.EOF Then Exit Sub
            numlinea = Data5.Recordset!numlinea
        Case 5 'Formacion Empresa
            If Data6.Recordset.EOF Then Exit Sub
            numlinea = Data6.Recordset!numlinea
    End Select
    
    ModificaLineas = 3 'Eliminar
    Sql = "�Seguro que desea eliminar la l�nea de " & TituloLinea & "?"
    Sql = Sql & vbCrLf & "Cod. Traba.: " & Format(Data1.Recordset!CodTraba, "000000")
    Sql = Sql & vbCrLf & "Nombre: " & Data1.Recordset!NomTraba
    Sql = Sql & vbCrLf & "Numlinea: " & numlinea
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        Sql = "Delete from " & NomTablaLineas & " where codtraba=" & Data1.Recordset!CodTraba
        Sql = Sql & " and numlinea=" & numlinea
        conn.Execute Sql

        ModificaLineas = 0
        Select Case Me.SSTab1.Tab
            Case 1: 'Estudios/Formacion
'                CancelaADODC (Data2)
                CargaGrid2 DataGrid1, Data2
'                CancelaADODC (Data2)
            Case 2: 'Habilidades
                CargaGrid2 DataGrid2, Data3
            Case 3: 'Experiencia Laboral
                CargaGrid2 DataGrid3, Data4
            Case 4 'Formacion Realizada
                CargaGrid2 DataGrid4, Data5
            Case 5 'Formacion Empresa
                CargaGrid2 DataGrid5, Data6
        End Select
'        CancelaADODC
    End If
    PonerFocoBtn Me.cmdRegresar
    
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar Lineas Trabajador", Err.Description
End Sub


Private Sub cmdRegresar_Click()
'Este es el boton Cabecera tambien
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        
    Else 'Se llama desde alg�n Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ning�n registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo = 1 Then PonerFoco Text1(0)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del form
    Me.Icon = frmPpal.Icon
    
    'Icono de imagen de e-mail
    Me.ImgMail(0).Picture = frmPpal.imgListComun.ListImages(20).Picture

    'ICONITOS DE LA BARRA
    btnAnyadir = 5
    btnPrimero = 19
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Bot�n Buscar
        .Buttons(2).Image = 2   'Bot�n Todos
        .Buttons(5).Image = 3   'Insertar Nuevo
        .Buttons(6).Image = 4   'Modificar
        .Buttons(7).Image = 5   'Borrar
        .Buttons(10).Image = 25 'Estudios/Formacion
        .Buttons(11).Image = 27 'Habilidades
        .Buttons(12).Image = 37 'Experiencia Laboral
        .Buttons(13).Image = 28 'Formacion Realizada
        .Buttons(14).Image = 29 'Formacion Empresa
        
        .Buttons(16).Image = 16  'Salir
        .Buttons(17).Image = 15  'Salir
        .Buttons(btnPrimero).Image = 6  'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 '�ltimo
    End With
    
    Me.SSTab1.Tab = 0
      
    LimpiarCampos   'Limpia los campos TextBox
    VieneDeBuscar = False
    PrimeraVez = True
         
    '## A mano
    NombreTabla = "straba"
    Ordenacion = " ORDER BY codtraba"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    Data1.RecordSource = "Select * from " & NombreTabla & " where codtraba=-1"
    Data1.Refresh
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1
    End If
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    Text1(24).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") 'Cod Almac
    Text2(24).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom Almac
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim cadB As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        If Me.imgFecha(0).Tag = 0 Then
            Screen.MousePointer = vbHourglass
            cadB = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
            PonerCadenaBusqueda
        Else
            'Centro de coste
            Text1(10).Text = RecuperaValor(CadenaDevuelta, 1)
            Text2(10).Text = RecuperaValor(CadenaDevuelta, 2)
        End If
    End If
    Screen.MousePointer = vbDefault
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


Private Sub frmF_Selec(vFecha As Date) 'Calendario Fechas
Dim indice As Byte
    indice = Val(imgFecha(0).Tag) + 11
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'CPostal
            Set frmCP = New frmCPostal
            frmCP.DatosADevolverBusqueda = "0"
            frmCP.Show vbModal
            Set frmCP = Nothing
            indice = 4
            VieneDeBuscar = True
            
        Case 1 'Almacen por defecto del trabajador
            Set frmAlm = New frmAlmAlPropios
            frmAlm.DatosADevolverBusqueda = "0"
            frmAlm.Show vbModal
            Set frmAlm = Nothing
            indice = 24
            
        Case 2 'Centros de coste de la conta
            Me.imgFecha(0).Tag = 10
               
          
        
            Screen.MousePointer = vbHourglass
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
            
               
               
    End Select
    PonerFoco Text1(indice)
    Screen.MousePointer = vbDefault
End Sub


Private Sub imgFecha_Click(Index As Integer) 'Abre calendario Fechas
Dim indice As Byte

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
   
   Set frmF = New frmCal
   frmF.Fecha = Now
   Me.imgFecha(0).Tag = Index
   indice = Index + 11
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.Fecha = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub

Private Sub ImgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Index = 0 Then
        dirMail = Text1(9).Text
    End If
    If LanzaMailGnral(dirMail) Then Espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnBuscar_Click()
    Me.SSTab1.Tab = 0
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de trabajadores
         BotonEliminarLinea
    Else   'Eliminar Trabajador
         BotonEliminar
    End If
End Sub

Private Sub mnEstudios_Click()
'Abre Mantenimiento de lineas  Estudios/Formacion
    BotonMtoLineas 1, "Estudios/Formacion"
    NomTablaLineas = "strab1"
End Sub

Private Sub mnExperiencia_Click()
'Abre Mantenimiento de lineas Experiencia Laboral
    BotonMtoLineas 3, "Experiencia Laboral"
    NomTablaLineas = "strab3"
End Sub

Private Sub mnFormEmpresa_Click()
'Abre Mantenimiento de lineas Formacion Empresa
    BotonMtoLineas 5, "Formaci�n Empresa"
    NomTablaLineas = "strab5"
End Sub

Private Sub mnFormRealizada_Click()
'Abre Mantenimiento de lineas Formacion Realizada
    BotonMtoLineas 4, "Formaci�n Realizada"
    NomTablaLineas = "strab4"
End Sub

Private Sub mnHabilidades_Click()
'Abre Mantenimiento de lineas Habilidades
    BotonMtoLineas 2, "Habilidades"
    NomTablaLineas = "strab2"
End Sub

Private Sub mnModificar_Click()
    If Modo = 5 Then 'Modificar lineas
         BotonModificarLinea
    Else   'Modificar Trabajador
         Me.SSTab1.Tab = 0
         If BLOQUEADesdeFormulario(Me) Then BotonModificar
    End If
End Sub

Private Sub mnNuevo_Click()
    If Modo = 5 Then 'A�adir lineas
         BotonAnyadirLinea
    Else 'A�adir Trabajador
         Me.SSTab1.Tab = 0
         BotonAnyadir
    End If
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub


'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Index = 0 And KeyCode = 38 Then Exit Sub
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
Dim devuelve As String
Dim cta As String
Dim CC As String
    
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si queremos hacer algo ..
    Select Case Index
        Case 0 'Cod. Trabajador
            If PonerFormatoEntero(Text1(Index)) Then
                'Comprobar si ya existe el cod de trabajador en la tabla
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
            
        Case 3 'CPostal
             If Not VieneDeBuscar Then
                Text1(Index + 1).Text = ObtenerPoblacion(Text1(Index).Text, devuelve)
                Text1(Index + 2).Text = devuelve
            End If
            VieneDeBuscar = False
            
        Case 6 'NIF
            Text1(Index).Text = UCase(Text1(Index).Text)
            ValidarNIF Text1(Index).Text
            
            
        Case 10
            ' ---- [19/10/2009] [LAURA] : A�adir funcion generica de ccoste
            Me.Text2(Index).Text = PonerNombreCCoste(Me.Text1(Index))
'            PonceCentroCoste
            
        Case 11, 12, 13 'Fecha Nacimiento, Fecha alta, Fecha baja
            'Si no es modo de Busqueda poner el formato
             If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index)
             
        Case 24 'Cod almacen
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), conAri, "salmpr", "nomalmac", "codalmac")
            Else
                Text2(Index).Text = ""
            End If
            
        Case 15, 16, 19, 20 'cod. banco, cod. sucursal
            PonerFormatoEntero Text1(Index)
            
        Case 25, 26 ' codigo de iban
            Text1(Index).Text = UCase(Text1(Index).Text)
            
    End Select
    
    '[Monica]: calculo del iban si no lo ponen
    If Index = 15 Or Index = 16 Or Index = 17 Or Index = 18 Then
        If Text1(15).Text <> "" And Text1(16).Text <> "" And Text1(17).Text <> "" And Text1(18).Text <> "" Then
            
            cta = Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "00") & Format(Text1(18).Text, "0000000000")
            If Len(cta) = 20 Then
                If Text1(25).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(25).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(25).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(25).Text, 3) <> cta Then
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If
            
    If Index = 19 Or Index = 20 Or Index = 21 Or Index = 22 Then
        If Text1(19).Text <> "" And Text1(20).Text <> "" And Text1(21).Text <> "" And Text1(22).Text <> "" Then
            cta = Format(Text1(19).Text, "0000") & Format(Text1(20).Text, "0000") & Format(Text1(21).Text, "00") & Format(Text1(22).Text, "0000000000")
            If Len(cta) = 20 Then
                If Text1(26).Text = "" Then
                    'NO ha puesto IBAN
                    If DevuelveIBAN2("ES", cta, cta) Then Text1(26).Text = "ES" & cta
                Else
                    CC = CStr(Mid(Text1(26).Text, 1, 2))
                    If DevuelveIBAN2(CStr(CC), cta, cta) Then
                        If Mid(Text1(26).Text, 3) <> cta Then
                            MsgBox "Codigo IBAN distinto del calculado [" & CC & cta & "]", vbExclamation
                        End If
                    End If
                End If
            End If
        End If
    End If
            
End Sub

' ---- [19/10/2009] [LAURA] : A�adir funcion generica de ccoste
'Private Sub PonceCentroCoste()
'Dim C As String
'    text1(10).Text = Trim(text1(10).Text)
'    C = ""
'    If text1(10).Text <> "" Then
'        C = PonerNombreDeCod(text1(10), conConta, "cabccost", "nomccost", "codccost")
'        If C = "" Then
'            MsgBox "No existe centro de coste", vbExclamation
'            text1(10).Text = ""
'        End If
'    End If
'    Text2(10).Text = C
'
'End Sub


Private Sub HacerBusqueda()
Dim cadB As String

    cadB = ObtenerBusqueda(Me, False)
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia cadB
    ElseIf cadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & cadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        PonerFoco Text1(0)
    End If
End Sub


Private Sub MandaBusquedaPrevia(cadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & ParaGrid(Text1(0), 14, "C�digo")
    cad = cad & ParaGrid(Text1(1), 65, "Nombre")
    cad = cad & ParaGrid(Text1(6), 18, "NIF")
'            cad = cad & ParaGrid(Text1(2), 40, "Nombre Comercial")
    Tabla = "straba"
    Titulo = "Trabajadores"
    Me.imgFecha(0).Tag = 0
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = cadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vselElem = 1
        frmB.vConexionGrid = conAri
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(0)
            Text1(0).BackColor = vbYellow
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
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



Private Sub PonerCamposLineas()
'Carga las Pesta�as con las tablas de lineas del Trabajador seleccionado para mostrar
Dim Sql As String
Dim vWhere As String
On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
   
    vWhere = " WHERE codtraba= " & Data1.Recordset!CodTraba
    'Estudios/Formacion - Datos de la tabla strab1
    Sql = "Select * from strab1 " & vWhere
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid1, Data2, Sql
    
    'Habilidades
    Sql = "Select * from strab2 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid2, Data3, Sql

    'Experiencia Laboral
    Sql = "Select * from strab3 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid3, Data4, Sql

    'Formacion Realizada
    Sql = "Select * from strab4 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid4, Data5, Sql

    'Formacion Empresa
    Sql = "Select * from strab5 " & vWhere 'where codtraba= " & Data1.Recordset!codtraba
    Sql = Sql & " order by numlinea"
    CargaGrid DataGrid5, Data6, Sql

    PrimeraVez = False
    Screen.MousePointer = vbDefault
    Exit Sub
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error Resume Next
    
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    Text2(24).Text = PonerNombreDeCod(Text1(24), conAri, "salmpr", "nomalmac")
    
    ' ---- [19/10/2009] [LAURA] : A�adir funcion generica de ccoste
'    PonceCentroCoste
    Me.Text2(10).Text = PonerNombreCCoste(Me.Text1(10))
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas asociadas al trabajador
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim i As Byte, NumReg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    'Visualizar el login solo si es administrador o root
    b = (vUsu.Nivel < 2)
    Me.Label1(25).visible = b
    Text1(23).visible = b

    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i



    'Actualiza Iconos Insertar,Modificar,Eliminar
    ActualizarToolbarGnral Me.Toolbar1, Modo, Kmodo, btnAnyadir
    
    
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
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2 And Modo <> 5
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    For i = 0 To Me.imgFecha.Count - 1
        Me.imgFecha(i).Enabled = b
    Next i
    
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Enabled = b
    Next i
    
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    'Poner el tama�o de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos

    '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opciones menu seg�n modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerModoOpcionesMenu(Modo)
'Activas unas Opciones de Menu y Toolbar seg�n el modo en que estemos
Dim b As Boolean
Dim i As Byte

    b = (Modo = 2 Or Modo = 5 Or Modo = 0 Or Modo = 1)
    'Insertar
    Toolbar1.Buttons(5).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 Or Modo = 5)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnEliminar.Enabled = b
    
    'Mantenimiento lineas
    b = (Modo = 2)
    For i = 10 To 14
        Toolbar1.Buttons(i).Enabled = b
    Next i
    Toolbar1.Buttons(16).Enabled = b Or Modo = 0
    Me.mnEstudios.Enabled = b
    Me.mnExperiencia.Enabled = b
    Me.mnFormEmpresa.Enabled = b
    Me.mnFormRealizada.Enabled = b
    Me.mnHabilidades.Enabled = b
    
    '------------------------------------------
    b = (Modo >= 3)
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b
    Me.mnBuscar.Enabled = Not b
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b
    Me.mnVerTodos.Enabled = Not b
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
  PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim cta As String
Dim cadMen As String

On Error GoTo EDatosOK

    DatosOk = False
    b = True
    b = CompForm(Me, 1) 'Comprobar formato datos ok
    If Not b Then Exit Function
'[Monica]22/11/2013: iban
'    Comprueba_CuentaBan (Text1(31).Text & Text1(32).Text & Text1(33).Text & Text1(34).Text)
    If b And (Modo = 3 Or Modo = 4) Then
        
        
        '[Monica]22/08/2013: a�adida la comprobacion de que la cuenta contable sea correcta
        If Text1(15).Text = "" Or Text1(16).Text = "" Or Text1(17).Text = "" Or Text1(18).Text = "" Then
            '[Monica]20/11/2013: a�adido el codigo de iban
            Text1(25).Text = ""
            Text1(15).Text = ""
            Text1(16).Text = ""
            Text1(17).Text = ""
            Text1(18).Text = ""
        Else
            cta = Format(Text1(15).Text, "0000") & Format(Text1(16).Text, "0000") & Format(Text1(17).Text, "00") & Format(Text1(18).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El trabajador no tiene asignada cuenta bancaria de n�mina."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria de n�mina del trabajador no es correcta. � Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(15)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(25).Text <> "" Then BuscaChekc = Mid(Text1(25).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(25).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(25).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(25).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(25).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(25)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
        '[Monica]22/08/2013: a�adida la comprobacion de que la cuenta contable sea correcta
        If Text1(19).Text = "" Or Text1(20).Text = "" Or Text1(21).Text = "" Or Text1(22).Text = "" Then
            '[Monica]20/11/2013: a�adido el codigo de iban
            Text1(26).Text = ""
            Text1(19).Text = ""
            Text1(20).Text = ""
            Text1(21).Text = ""
            Text1(22).Text = ""
        Else
            cta = Format(Text1(19).Text, "0000") & Format(Text1(20).Text, "0000") & Format(Text1(21).Text, "00") & Format(Text1(22).Text, "0000000000")
            If Val(ComprobarCero(cta)) = 0 Then
                cadMen = "El trabajador no tiene asignada cuenta bancaria de gastos."
                MsgBox cadMen, vbExclamation
            End If
            If Not Comprueba_CC(cta) Then
                cadMen = "La cuenta bancaria de gastos del trabajador no es correcta. � Desea continuar ?."
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    b = True
                Else
                    PonerFoco Text1(19)
                    b = False
                End If
            Else
                BuscaChekc = ""
                If Me.Text1(26).Text <> "" Then BuscaChekc = Mid(Text1(26).Text, 1, 2)
                    
                If DevuelveIBAN2(BuscaChekc, cta, cta) Then
                    If Me.Text1(26).Text = "" Then
                        If MsgBox("Poner IBAN ?", vbQuestion + vbYesNo) = vbYes Then Me.Text1(26).Text = BuscaChekc & cta
                    Else
                        If Mid(Text1(26).Text, 3) <> cta Then
                            cta = "Calculado : " & BuscaChekc & cta
                            cta = "Introducido: " & Me.Text1(26).Text & vbCrLf & cta & vbCrLf
                            cta = "Error en codigo IBAN" & vbCrLf & cta & "Continuar?"
                            If MsgBox(cta, vbQuestion + vbYesNo) = vbNo Then
                                PonerFoco Text1(19)
                                b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    
    
    
    
    End If
          
    DatosOk = b
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim b As Boolean
On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    b = True

    Select Case Me.SSTab1.Tab
        Case 1 'Estudios/Formacion
            If (Not Trim(txtAux1(0).Text) <> "") And (Not Trim(txtAux1(1).Text) <> "") Then
                MsgBox "Los campos Periodo y Formaci�n no pueden ser nulos", vbExclamation
                b = False
            End If
        Case 2 'Habilidades
            If Trim(txtAux2.Text) = "" Then
                MsgBox "El campo Habilidad no puede ser nulo", vbExclamation
                b = False
            End If
        Case 3 'Experiencial Laboral
            If (Not Trim(txtAux3(0).Text) <> "") And (Not Trim(txtAux3(1).Text) <> "") Then
                MsgBox "Los campos Periodo y Experiencia no pueden ser nulos", vbExclamation
                b = False
            End If
        Case 4 'Formacion Realizada
            If (Not Trim(txtAux4(0).Text) <> "") And (Not Trim(TxtAux5(1).Text) <> "") Then
                MsgBox "Los campos Fecha Formaci�n y Fecha Evaluaci�n no pueden ser nulos", vbExclamation
                b = False
            End If
        Case 5 'Formacion Empresa
            If Trim(TxtAux5(0).Text) = "" Then
                MsgBox "El campo Fecha Formaci�n no puede ser nulo", vbExclamation
                b = False
            End If
    End Select
    
    DatosOkLinea = b
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 1  'Buscar
           mnBuscar_Click
        Case 2  'Todos
            BotonVerTodos
            
        Case 5  'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7  'Borrar
            mnEliminar_Click
            
        Case 10  'Estudios/Formacion
            mnEstudios_Click
        Case 11  'Habilidades
            mnHabilidades_Click
        Case 12  'Experiencia Laboral
            mnExperiencia_Click
        Case 13 'Formacion Realizada
            mnFormRealizada_Click
        Case 14  'Formacion Empresa
            mnFormEmpresa_Click
            
        Case 16
            frmListado2.Opcion = 17
            frmListado2.Show vbModal
            
        Case 17    'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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
   
    
Private Function InsertarLinea() As Boolean
Dim Sql As String
Dim vWhere As String
Dim numF As String
On Error GoTo EInsertarLinea

    InsertarLinea = False
    Sql = ""
    If DatosOkLinea Then
          vWhere = "codtraba=" & Val(Text1(0).Text)
          numF = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
          Select Case Me.SSTab1.Tab
             Case 1 'Estudios/Formacion
                 Sql = "INSERT INTO strab1 "
                 Sql = Sql & "(codtraba, numlinea, periodos, formacio) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
                 Sql = Sql & DBSet(txtAux1(0).Text, "T") & "," & DBSet(txtAux1(1).Text, "T") & ")"
            Case 2 'Habilidades
                 Sql = "INSERT INTO strab2 "
                 Sql = Sql & "(codtraba, numlinea, habilida) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
                 Sql = Sql & DBSet(txtAux2.Text, "T") & ")"
            Case 3 'Experiencia Laboral
                 Sql = "INSERT INTO strab3 "
                 Sql = Sql & "(codtraba, numlinea, periodo1, experien) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ","
                 Sql = Sql & DBSet(txtAux3(0).Text, "T") & ", " & DBSet(txtAux3(1).Text, "T") & ")"
            Case 4 'Formacion Realizada
                 Sql = "INSERT INTO strab4 "
                 Sql = Sql & "(codtraba, numlinea, fecforma, fecevalu, formaci1, codcentr, evaluaci) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ",'"
                 Sql = Sql & Format(txtAux4(0).Text, FormatoFecha) & "', '" & Format(txtAux4(1).Text, FormatoFecha) & "', "
                 Sql = Sql & DBSet(txtAux4(2).Text, "T") & ", " & DBSet(txtAux4(3).Text, "T") & ", " & DBSet(txtAux4(4).Text, "T") & ")"
            Case 5 'Formacion Empresa
                 Sql = "INSERT INTO strab5 "
                 Sql = Sql & "(codtraba, numlinea, fecforma, formaci2, resforma) "
                 Sql = Sql & "VALUES (" & Val(Text1(0).Text) & ", " & numF & ",'"
                 Sql = Sql & Format(TxtAux5(0).Text, FormatoFecha) & "', " & DBSet(TxtAux5(1).Text, "T") & ", " & DBSet(TxtAux5(2).Text, "T") & ")"
          End Select
     End If
    
    If Sql <> "" Then
        conn.Execute Sql
        InsertarLinea = True
    End If
    Exit Function
EInsertarLinea:
    MuestraError Err.Number, "Insertar Lineas Trabajador" & vbCrLf & Err.Description
End Function


Private Function ModificarLinea() As Boolean
Dim Sql As String
Dim vWhere As String
On Error GoTo EModificarLinea

    ModificarLinea = False
    Sql = ""
    If DatosOkLinea Then
         vWhere = "codtraba=" & Val(Text1(0).Text)
         Select Case Me.SSTab1.Tab
            Case 1 'Estudios/Formacion
                Sql = "UPDATE strab1 Set periodos = " & DBSet(txtAux1(0).Text, "T")
                Sql = Sql & ", formacio = " & DBSet(txtAux1(1).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data2.Recordset!numlinea
            Case 2 'Habilidades
                Sql = "UPDATE strab2 Set habilida = " & DBSet(txtAux2.Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data3.Recordset!numlinea
            Case 3 'Experiencia Laboral
                Sql = "UPDATE strab3 Set periodo1 = " & DBSet(txtAux3(0).Text, "T") & ", "
                Sql = Sql & "experien=" & DBSet(txtAux3(1).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data4.Recordset!numlinea
            Case 4 'Formacion Realizada
                Sql = "UPDATE strab4 Set fecforma = " & DBSet(txtAux4(0).Text, "F") & ", fecevalu="
                Sql = Sql & DBSet(txtAux4(1).Text, "F") & ", formaci1=" & DBSet(txtAux4(2).Text, "T") & ", codcentr="
                Sql = Sql & DBSet(txtAux4(3).Text, "T") & ", evaluaci=" & DBSet(txtAux4(4).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data5.Recordset!numlinea
            Case 5 'Formacion Empresa
                Sql = "UPDATE strab5 Set fecforma = " & DBSet(TxtAux5(0).Text, "F") & ", formaci2="
                Sql = Sql & DBSet(TxtAux5(1).Text, "T") & ", resforma=" & DBSet(TxtAux5(2).Text, "T")
                Sql = Sql & " WHERE " & vWhere & " AND numlinea=" & Data6.Recordset!numlinea
        End Select
    End If

    If Sql <> "" Then
        conn.Execute Sql
        ModificarLinea = True
    End If
    Exit Function
EModificarLinea:
    MuestraError Err.Number, "Modificar Lineas Trabajador" & vbCrLf & Err.Description
End Function


Private Sub PonerBotonCabecera(b As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
On Error Resume Next
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdRegresar.visible = b
    Me.cmdRegresar.Caption = "Cabecera"
    If b Then
        Me.lblIndicador.Caption = "L�neas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, Sql As String)
On Error GoTo ECargaGrid

    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    vDataGrid.RowHeight = 320
     
    CargaGrid2 vDataGrid, vData
    vDataGrid.Enabled = (Modo = 0 Or Modo = 2)
    vDataGrid.ScrollBars = dbgAutomatic
    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim i As Integer

On Error GoTo ECargaGrid

    vData.Refresh
    
    vDataGrid.Columns(0).visible = False 'codtraba
    vDataGrid.Columns(1).visible = False 'numlinea

    Select Case vDataGrid.Name
        Case "DataGrid1" 'Estudios / Formacion
                vDataGrid.Columns(2).Caption = "Periodo"
                vDataGrid.Columns(2).Width = 2100
                vDataGrid.Columns(3).visible = True
                vDataGrid.Columns(3).Caption = "Formaci�n"
                vDataGrid.Columns(3).Width = 6450
        Case "DataGrid2" 'Habilidades
                vDataGrid.Columns(2).Caption = "Habilidades"
                vDataGrid.Columns(2).Width = 7250
        Case "DataGrid3" 'Experiencia Laboral
                vDataGrid.Columns(2).Caption = "Periodo"
                vDataGrid.Columns(2).Width = 2100
                vDataGrid.Columns(3).visible = True
                vDataGrid.Columns(3).Caption = "Experiencia"
                vDataGrid.Columns(3).Width = 6450
        Case "DataGrid4" 'Formacion Realizada
                vDataGrid.Columns(2).Caption = "Fecha Formac."
                vDataGrid.Columns(2).Width = 1450
                vDataGrid.Columns(3).Caption = "Fecha Eval."
                vDataGrid.Columns(3).Width = 1450
                vDataGrid.Columns(4).Caption = "Formaci�n"
                vDataGrid.Columns(4).Width = 4000
                vDataGrid.Columns(5).Caption = "Centro"
                vDataGrid.Columns(5).Width = 1670
                vDataGrid.Columns(6).Caption = "Evaluaci�n"
                vDataGrid.Columns(6).Width = 1160
        Case "DataGrid5" 'Formacion Empresa
                vDataGrid.Columns(2).Caption = "Fecha Formac."
                vDataGrid.Columns(2).Width = 1500
                vDataGrid.Columns(3).Caption = "Formaci�n"
                vDataGrid.Columns(3).Width = 4670
                vDataGrid.Columns(4).Caption = "Resultado"
                vDataGrid.Columns(4).Width = 1900
    End Select

    vDataGrid.Enabled = (Modo = 0) Or (Modo = 2) Or (Modo = 5 And ModificaLineas = 0)
    For i = 0 To vDataGrid.Columns.Count - 1
        vDataGrid.Columns(i).Locked = True
        vDataGrid.Columns(i).AllowSizing = False
    Next i

    Exit Sub
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux1(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux1.Count - 1 'TextBox
            txtAux1(i).Top = 290
            txtAux1(i).visible = visible
        Next i
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid1
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = ""
                BloquearTxt txtAux1(i), False
            Next i
        Else
            For i = 0 To txtAux1.Count - 1
                txtAux1(i).Text = DataGrid1.Columns(i + 2).Text
                BloquearTxt txtAux1(i), False
            Next i
        End If


        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid1, 8)
        
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).Top = alto
            txtAux1(i).Height = DataGrid1.RowHeight
        Next i
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Periodo
        txtAux1(0).Left = DataGrid1.Left + 320
        txtAux1(0).Width = DataGrid1.Columns(2).Width - 20
        'Formacion
        txtAux1(1).Left = txtAux1(0).Left + txtAux1(0).Width + 20
        txtAux1(1).Width = DataGrid1.Columns(3).Width - 20
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux1.Count - 1
            txtAux1(i).visible = visible
        Next i
    End If
End Sub


'Esta funcion sustituye a LlamaLineas
Private Sub CargaTxtAux2(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
            txtAux2.Top = 290
            txtAux2.visible = visible
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid2
            txtAux2.Text = ""
            BloquearTxt txtAux2, False
        Else
            txtAux2.Text = DataGrid2.Columns(2).Text
            BloquearTxt txtAux2, False
        End If


        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid2, 8)
        
        txtAux2.Top = alto
        txtAux2.Height = DataGrid2.RowHeight
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Habilidad
        txtAux2.Left = DataGrid2.Left + 320
        txtAux2.Width = DataGrid2.Columns(2).Width - 20
            
        'Los ponemos Visibles o No
        '--------------------------
        txtAux2.visible = visible
    End If
End Sub


Private Sub CargaTxtAux3(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte
    
    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux3.Count - 1 'TextBox
            txtAux3(i).Top = 290
            txtAux3(i).visible = visible
        Next i
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid3
            For i = 0 To txtAux3.Count - 1
                txtAux3(i).Text = ""
                BloquearTxt txtAux3(i), False
            Next i
        Else
            For i = 0 To txtAux3.Count - 1
                txtAux3(i).Text = DataGrid3.Columns(i + 2).Text
                BloquearTxt txtAux3(i), False
            Next i
        End If


        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid3, 8)
        
        For i = 0 To txtAux3.Count - 1
            txtAux3(i).Top = alto
            txtAux3(i).Height = DataGrid3.RowHeight
        Next i
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Periodo
        txtAux3(0).Left = DataGrid3.Left + 320
        txtAux3(0).Width = DataGrid3.Columns(2).Width - 20
        'Experiencia
        txtAux3(1).Left = txtAux3(0).Left + txtAux3(0).Width + 20
        txtAux3(1).Width = DataGrid3.Columns(3).Width - 20
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux3.Count - 1
            txtAux3(i).visible = visible
        Next i
    End If
End Sub


Private Sub CargaTxtAux4(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To txtAux4.Count - 1 'TextBox
            txtAux4(i).Top = 290
            txtAux4(i).visible = visible
        Next i
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid4
            For i = 0 To txtAux4.Count - 1
                txtAux4(i).Text = ""
                BloquearTxt txtAux4(i), False
            Next i
        Else
            For i = 0 To txtAux4.Count - 1
                txtAux4(i).Text = DataGrid4.Columns(i + 2).Text
                BloquearTxt txtAux4(i), False
            Next i
        End If


        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid4, 8)
        
        For i = 0 To txtAux4.Count - 1
            txtAux4(i).Top = alto
            txtAux4(i).Height = DataGrid4.RowHeight
        Next i
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Fecha Formacion
        txtAux4(0).Left = DataGrid4.Left + 320
        txtAux4(0).Width = DataGrid4.Columns(2).Width - 20
        'Fecha Evaluacion
        For i = 1 To 4
            txtAux4(i).Left = txtAux4(i - 1).Left + txtAux4(i - 1).Width + 20
            txtAux4(i).Width = DataGrid4.Columns(i + 2).Width - 20
        Next i
        
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To txtAux4.Count - 1
            txtAux4(i).visible = visible
        Next i
    End If
End Sub


Private Sub CargaTxtAux5(visible As Boolean, limpiar As Boolean)
'IN: visible: si es true ponerlos visibles en la posici�n adecuada
'    limpiar: si es true vaciar los txtAux
Dim alto As Single
Dim i As Byte

'Formacion Empresa

    If Not visible Then
        'Fijamos el alto (ponerlo en la parte inferior del form)
        For i = 0 To TxtAux5.Count - 1 'TextBox
            TxtAux5(i).Top = 290
            TxtAux5(i).visible = visible
        Next i
    Else
        If limpiar Then 'Vaciar los textBox (Vamos a Insertar)
            DeseleccionaGrid DataGrid5
            For i = 0 To TxtAux5.Count - 1
                TxtAux5(i).Text = ""
                BloquearTxt TxtAux5(i), False
            Next i
        Else
            For i = 0 To TxtAux5.Count - 1
                TxtAux5(i).Text = DataGrid5.Columns(i + 2).Text
                BloquearTxt TxtAux5(i), False
            Next i
        End If


        'Fijamos altura(Height) y posici�n Top
        '-------------------------------
        alto = ObtenerAlto(DataGrid5, 8)
        
        For i = 0 To TxtAux5.Count - 1
            TxtAux5(i).Top = alto
            TxtAux5(i).Height = DataGrid5.RowHeight
        Next i
        
        'Fijamos anchura y posicion Left
        '--------------------------------
        'Fecha Formacion
        TxtAux5(0).Left = DataGrid5.Left + 320
        TxtAux5(0).Width = DataGrid5.Columns(2).Width - 20
        'Formacion
        TxtAux5(1).Left = TxtAux5(0).Left + TxtAux5(0).Width + 20
        TxtAux5(1).Width = DataGrid5.Columns(3).Width - 20
        'Resultado
        TxtAux5(2).Left = TxtAux5(1).Left + TxtAux5(1).Width + 20
        TxtAux5(2).Width = DataGrid5.Columns(4).Width - 20
                
        'Los ponemos Visibles o No
        '--------------------------
        For i = 0 To TxtAux5.Count - 1
            TxtAux5(i).visible = visible
        Next i
    End If
End Sub


Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFoco txtAux1(Index), Modo
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
      If Not (Index = 0 And KeyCode = 38) Then
            KEYdown KeyCode
      End If
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub BotonMtoLineas(numTab As Integer, cad As String)
        Me.SSTab1.Tab = numTab
        NumTabMto = numTab
        TituloLinea = cad
        PonerModo 5
        PonerBotonCabecera True
End Sub


Private Sub TxtAux1_LostFocus(Index As Integer)
    If Index = 1 And Me.SSTab1.Tab = 1 Then
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub txtAux2_GotFocus()
    ConseguirFoco txtAux2, Modo
End Sub

Private Sub txtAux2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then
           KEYdown KeyCode
    End If
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
     If Index = 1 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub


Private Sub TxtAux4_GotFocus(Index As Integer)
    ConseguirFoco txtAux4(Index), Modo
End Sub


Private Sub TxtAux4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then
           KEYdown KeyCode
    End If
End Sub

Private Sub TxtAux4_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 4 And KeyAscii = 13 Then
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub TxtAux4_LostFocus(Index As Integer)
    If Trim(txtAux4(Index).Text) = "" Then Exit Sub
    
    Select Case Index
        Case 0, 1
             PonerFormatoFecha txtAux4(Index)
    End Select
End Sub


Private Sub TxtAux5_GotFocus(Index As Integer)
    ConseguirFoco TxtAux5(Index), Modo
End Sub

Private Sub TxtAux5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    If Not (Index = 0 And KeyCode = 38) Then
        KEYdown KeyCode
    End If
End Sub

Private Sub TxtAux5_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub TxtAux5_LostFocus(Index As Integer)
    If Trim(TxtAux5(Index).Text) = "" Then Exit Sub
    Select Case Index
        Case 0
            PonerFormatoFecha TxtAux5(Index)
        Case 2
'            If Me.ActiveControl.TabIndex <> 72 Then PonerFocoBtn Me.cmdAceptar
            PonerFocoBtn Me.cmdAceptar
    End Select
End Sub


Private Function Eliminar() As Boolean
Dim Sql As String
On Error GoTo FinEliminar

        conn.BeginTrans
        Sql = " WHERE  codtraba=" & Data1.Recordset!CodTraba

        'Lineas Estudios/Formacion
        conn.Execute "Delete from strab1 " & Sql
        'Lineas Habilidades
        conn.Execute "Delete from strab2 " & Sql
        'Lineas Experiencia Laboral
        conn.Execute "Delete from strab3 " & Sql
        'Lineas Formacion Realizada
        conn.Execute "Delete from strab4 " & Sql
        'Lineas Experiencia Empresa
        conn.Execute "Delete from strab5 " & Sql
        'Cabeceras
        conn.Execute "Delete from straba " & Sql

FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function



Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ning�n registro
Dim cad As String
On Error Resume Next

    cad = "Select * from strab1 where codtraba= -1"
    CargaGrid DataGrid1, Data2, cad
    cad = "Select * from strab2 where codtraba= -1"
    CargaGrid DataGrid2, Data3, cad
    cad = "Select * from strab3 where codtraba= -1"
    CargaGrid DataGrid3, Data4, cad
    cad = "Select * from strab4 where codtraba= -1"
    CargaGrid DataGrid4, Data5, cad
    cad = "Select * from strab5 where codtraba= -1"
    CargaGrid DataGrid5, Data6, cad
    
    PrimeraVez = False
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    cad = "(codtraba=" & Text1(0).Text & ")"
    If SituarData(Data1, cad, Indicador) Then
       PonerModo 2
       lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       'Poner los grid sin apuntar a nada
       LimpiarDataGrids
       PonerModo 0
    End If
End Sub



