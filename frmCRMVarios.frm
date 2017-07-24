VERSION 5.00
Begin VB.Form frmCRMVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CRM"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameGenerar 
      Height          =   6735
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7455
      Begin VB.TextBox txtDescAccion 
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txtAccion 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdGeneAcciones 
         Caption         =   "Generar"
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
         Left            =   4920
         TabIndex        =   11
         Top             =   6120
         Width           =   1135
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
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
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   6120
         Width           =   1095
      End
      Begin VB.TextBox txtDescNumero 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txtDescNumero 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   5760
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtDescNumero 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   5280
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.TextBox txtDescNumero 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   4920
         Width           =   2295
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.TextBox txtDescNumero 
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
         Index           =   1
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   4320
         Width           =   3495
      End
      Begin VB.TextBox txtDescNumero 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtNumero 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
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
         Index           =   1
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtCliente 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtDescClie 
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
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox txtDescTra 
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
         Left            =   2310
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1140
         Width           =   3735
      End
      Begin VB.TextBox txtTrab 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   1
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         BeginProperty Font 
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
         Left            =   1350
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   6120
         Width           =   1135
      End
      Begin VB.Image imgAccion 
         Height          =   240
         Index           =   0
         Left            =   1110
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Accion"
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
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   1560
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   7320
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Caption         =   "Generar entrada de acciones comerciales"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   6855
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
         Index           =   5
         Left            =   360
         TabIndex        =   36
         Top             =   5760
         Width           =   705
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
         Index           =   4
         Left            =   360
         TabIndex        =   35
         Top             =   6120
         Width           =   705
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
         Index           =   3
         Left            =   360
         TabIndex        =   34
         Top             =   4920
         Width           =   705
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
         Index           =   2
         Left            =   360
         TabIndex        =   33
         Top             =   5280
         Width           =   705
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
         Index           =   1
         Left            =   360
         TabIndex        =   32
         Top             =   3960
         Width           =   705
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
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   4320
         Width           =   705
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
         Index           =   58
         Left            =   360
         TabIndex        =   30
         Top             =   3000
         Width           =   705
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
         Index           =   59
         Left            =   360
         TabIndex        =   29
         Top             =   3360
         Width           =   705
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   5
         Left            =   1110
         Top             =   6120
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
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
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   5520
         Width           =   495
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   4
         Left            =   1110
         Top             =   5760
         Width           =   240
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   3
         Left            =   1110
         Top             =   5310
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Zona"
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
         Index           =   2
         Left            =   240
         TabIndex        =   24
         Top             =   4680
         Width           =   540
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   2
         Left            =   1110
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Agente"
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
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   3720
         Width           =   780
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   1
         Left            =   1110
         Top             =   4320
         Width           =   240
      End
      Begin VB.Image imgVarios 
         Height          =   240
         Index           =   0
         Left            =   1110
         Top             =   3960
         Width           =   240
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   1
         Left            =   1110
         Top             =   3360
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Técnico"
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
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1140
         Width           =   825
      End
      Begin VB.Image imgCliente 
         Height          =   240
         Index           =   0
         Left            =   1110
         Top             =   3000
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   8
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   765
      End
      Begin VB.Image imgTecnico 
         Height          =   240
         Index           =   0
         Left            =   1110
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1110
         Top             =   720
         Width           =   240
      End
      Begin VB.Label lblDpto 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Index           =   33
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCRMVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '   0.-  Generacion


Private WithEvents frmCli As frmFacClientes
Attribute frmCli.VB_VarHelpID = -1
Private WithEvents frmT As frmAdmTrabajadores
Attribute frmT.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmAcc As frmCRMtipos
Attribute frmAcc.VB_VarHelpID = -1

Dim IndiceImg As Integer
Dim miSQL As String
Dim Codigo As String


Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdGeneAcciones_Click()
    miSQL = ""
    If txtFecha(0).Text = "" Then miSQL = miSQL & "- Fecha debe tener valor" & vbCrLf
    If txtTrab(0).Text = "" Then miSQL = miSQL & "- Trabajador debe tener valor" & vbCrLf
    If txtAccion(0).Text = "" Then miSQL = miSQL & "- Indique la accion comercial" & vbCrLf
    
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = ""
    miSQL = ""
    
    'Desde hasta
    If txtCliente(0).Text <> "" Then miSQL = miSQL & " AND sclien.codclien >= " & txtCliente(0).Text
    If txtCliente(1).Text <> "" Then miSQL = miSQL & " AND sclien.codclien <= " & txtCliente(1).Text
    'Agente
    If txtNumero(0).Text <> "" Then miSQL = miSQL & " AND sclien.codagent >= " & txtNumero(0).Text
    If txtNumero(1).Text <> "" Then miSQL = miSQL & " AND sclien.codagent <= " & txtNumero(1).Text
    'ZOna
    If txtNumero(2).Text <> "" Then miSQL = miSQL & " AND sclien.codzonas >= " & txtNumero(2).Text
    If txtNumero(3).Text <> "" Then miSQL = miSQL & " AND sclien.codzonas <= " & txtNumero(3).Text
    'RUTA
    If txtNumero(4).Text <> "" Then miSQL = miSQL & " AND sclien.codrutas >= " & txtNumero(4).Text
    If txtNumero(5).Text <> "" Then miSQL = miSQL & " AND sclien.codrutas <= " & txtNumero(5).Text
    
    If miSQL <> "" Then miSQL = Mid(miSQL, 5) 'quito el primer AND
    
    If Not HayRegParaInforme("sclien", miSQL, True) Then
        MsgBox "No hay clientes con estos valores", vbExclamation
        Exit Sub
    End If
    
    CadenaDesdeOtroForm = miSQL
    If CadenaDesdeOtroForm <> "" Then CadenaDesdeOtroForm = " WHERE " & CadenaDesdeOtroForm
    
    frmVarios.Opcion = 4
    frmVarios.Show vbModal
    
    Screen.MousePointer = vbHourglass
    If CadenaDesdeOtroForm <> "" Then
        DoEvents
        GenerarEntradaMasivaAccionesComerciales
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
Dim i As Integer
    Me.Icon = frmppal.Icon
    FrameGenerar.visible = False
    limpiar Me
    
    If Opcion = 0 Then
        H = 6735
        W = 7455
        PonerFrameVisible FrameGenerar, H, W
        txtFecha(0).Text = Format(Now, "dd/mm/yyyy")
        
        
        txtTrab(0).Text = PonerTrabajadorConectado(miSQL)
        Me.txtDescTra(0).Text = miSQL
        miSQL = ""
    End If
    
    For i = 0 To Me.imgAccion.Count - 1
        imgAccion(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    For i = 0 To Me.imgFecha.Count - 1
        imgFecha(i).Picture = frmppal.imgIcoForms.ListImages(2).Picture
    Next
    For i = 0 To Me.imgCliente.Count - 1
        imgCliente(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    For i = 0 To Me.imgTecnico.Count - 1
        imgTecnico(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    For i = 0 To Me.imgVarios.Count - 1
        imgVarios(i).Picture = frmppal.imgIcoForms.ListImages(1).Picture
    Next
    
    Me.cmdCancelar(Opcion).Cancel = True
    Me.Height = H
    Me.Width = W
    
End Sub

Private Sub PonerFrameVisible(ByRef F As Frame, ByRef CH As Integer, CW As Integer)
    F.top = 0
    F.Left = 120
    F.visible = True
    CH = F.Height + 420
    CW = F.Width + 240
End Sub


Private Sub frmAcc_DatoSeleccionado(CadenaSeleccion As String)
    miSQL = CadenaSeleccion
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    miSQL = CadenaDevuelta
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Me.txtFecha(IndiceImg).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtCliente(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    Me.txtDescClie(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmT_DatoSeleccionado(CadenaSeleccion As String)
    Me.txtTrab(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 1)
    txtDescTra(IndiceImg).Text = RecuperaValor(CadenaSeleccion, 2)
 
End Sub

Private Sub imgAccion_Click(Index As Integer)
        miSQL = ""
        Set frmAcc = New frmCRMtipos
        frmAcc.DatosADevolverBusqueda = "0|1|"
        frmAcc.Show vbModal
        Set frmAcc = Nothing
        If miSQL <> "" Then
            'Por defecto
            'NO dejo que la accon sea del 1 al 20 ya que las reservamos para otros menesteres
            If Val(RecuperaValor(miSQL, 1)) <= 20 Then
                MsgBox "Codigos reservados para la aplicacion", vbExclamation
                
            Else
                txtAccion(Index).Text = RecuperaValor(miSQL, 1)  'Pongo EL ID
                txtDescAccion(Index).Text = RecuperaValor(miSQL, 2)
            End If
        End If
End Sub

Private Sub imgCliente_Click(Index As Integer)
    
    Screen.MousePointer = vbHourglass
    IndiceImg = Index
    Set frmCli = New frmFacClientes
    frmCli.DatosADevolverBusqueda = "0|1|"
    frmCli.Show vbModal
    Set frmCli = Nothing
End Sub


Private Sub imgFecha_Click(Index As Integer)
   IndiceImg = Index
   Set frmC = New frmCal
   frmC.Fecha = Now
   If txtFecha(IndiceImg).Text <> "" Then
        If IsDate(txtFecha(IndiceImg).Text) Then frmC.Fecha = CDate(txtFecha(IndiceImg).Text)
    End If
   frmC.Show vbModal
   Set frmC = Nothing
    
End Sub



Private Sub imgTecnico_Click(Index As Integer)
   
        Set frmT = New frmAdmTrabajadores
        frmT.DatosADevolverBusqueda = "0|1|"
        frmT.Show vbModal
        Set frmT = Nothing
End Sub

Private Sub imgVarios_Click(Index As Integer)
Dim campo As String

    miSQL = ""
    
    'EN codigo:
    'titulo|tabla|sql|
    
    Set frmB = New frmBuscaGrid
    Select Case Index
    Case 0, 1
        'AGENTE
        campo = "Cod.|sagent|codagent|N||20·"
        campo = campo & "Nombre|sagent|nomagent|T||40·"
        
        Codigo = "Agente|sagent||"
    Case 2, 3
        'ZONA
        campo = "Cod.|szonas|codzonas|N||20·"
        campo = campo & "Desc.|szonas|nomzonas|T||40·"
        
        Codigo = "ZONAS|szonas||"
    
    Case 4, 5
        'RUTA
        campo = "Cod.|srutas|codrutas|N||20·"
        campo = campo & "Desc.|srutas|nomrutas|T||40·"
        
        Codigo = "Rutas|srutas||"
    End Select
    frmB.vCampos = campo
    frmB.vTitulo = RecuperaValor(Codigo, 1)
    frmB.vTabla = RecuperaValor(Codigo, 2)
    frmB.vSQL = RecuperaValor(Codigo, 3)
    frmB.vCargaFrame = False
    frmB.vDevuelve = "0|1|"
    frmB.vselElem = 1
    frmB.vConexionGrid = 1  'ODBC Aritaxi
    frmB.Show vbModal
    Set frmB = Nothing
    Screen.MousePointer = vbDefault
            
    If miSQL <> "" Then
        txtNumero(Index).Text = RecuperaValor(miSQL, 1)
        txtDescNumero(Index) = RecuperaValor(miSQL, 2)
            
            
        miSQL = ""
    End If

End Sub

Private Sub txtAccion_GotFocus(Index As Integer)
    ConseguirFoco txtAccion(Index), 3
End Sub

Private Sub txtAccion_KeyPress(Index As Integer, KeyAscii As Integer)
     KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtAccion_LostFocus(Index As Integer)
     miSQL = ""
    txtAccion(Index).Text = Trim(txtAccion(Index).Text)
    If txtAccion(Index).Text <> "" Then
        If Not IsNumeric(txtAccion(Index).Text) Then
            MsgBox "Campo accion debe ser numérico", vbExclamation
            txtAccion(Index).Text = ""
            PonerFoco txtAccion(Index)
        Else
            If Val(txtAccion(Index).Text) < 21 Then
                MsgBox "Las 20 primeras se las reserva la aplicacion", vbExclamation
                miSQL = ""
            Else
                miSQL = DevuelveDesdeBD(conAri, "denominacion", "scrmtipo", "codigo", txtAccion(Index).Text, "N")
                If miSQL = "" Then MsgBox "No existe la accion comercial : " & txtAccion(Index).Text, vbExclamation
            End If
            If miSQL = "" Then
                txtAccion(Index).Text = ""
                PonerFoco txtAccion(Index)
            End If
        End If
    End If
    Me.txtDescAccion(Index).Text = miSQL
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
    ConseguirFoco txtCliente(Index), 3
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)

    
    miSQL = ""
    txtCliente(Index).Text = Trim(txtCliente(Index).Text)
    If txtCliente(Index).Text <> "" Then
        If Not IsNumeric(txtCliente(Index).Text) Then
            MsgBox "Campo codigo cliente debe ser numérico", vbExclamation
            txtCliente(Index).Text = ""
            PonerFoco txtCliente(Index)
        Else
            miSQL = DevuelveDesdeBD(conAri, "nomclien", "sclien", "codclien", txtCliente(Index).Text, "N")
            If miSQL = "" Then
                MsgBox "No existe el cliente : " & txtCliente(Index).Text, vbExclamation
            End If
        End If
    End If
    Me.txtDescClie(Index).Text = miSQL
    
    
    
End Sub

Private Sub txtFecha_GotFocus(Index As Integer)
    ConseguirFoco txtFecha(Index), 3
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
Dim T As String
    txtFecha(Index).Text = Trim(txtFecha(Index).Text)
    If txtFecha(Index).Text <> "" Then
        T = txtFecha(Index).Text
        If EsFechaOK(T) Then
            txtFecha(Index).Text = T
        Else
            MsgBox "Fecha con formato incorrecto: " & txtFecha(Index).Text, vbExclamation
            PonerFoco txtFecha(Index)
        End If
    End If
    
End Sub





Private Sub txtNumero_GotFocus(Index As Integer)
     ConseguirFoco txtNumero(Index), 3
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtNumero_LostFocus(Index As Integer)

    txtNumero(Index).Text = Trim(txtNumero(Index).Text)
    miSQL = ""
    If txtNumero(Index).Text <> "" Then
        If Not IsNumeric(txtNumero(Index).Text) Then
            MsgBox "Campo debe ser numérico: " & txtNumero(Index).Text, vbExclamation
            txtNumero(Index).Text = ""
            PonerFoco txtNumero(Index)
        Else
            'Segun sea
            Select Case Index
            Case 0, 1
                'AGENTE
                miSQL = DevuelveDesdeBD(conAri, "nomagent", "sagent", "codagent", txtNumero(Index).Text, "N")
            Case 2, 3
                'ZONA
                miSQL = DevuelveDesdeBD(conAri, "nomzonas", "szonas", "codzonas", txtNumero(Index).Text, "N")
            Case 4, 5
                'RUTA
                miSQL = DevuelveDesdeBD(conAri, "nomrutas", "srutas", "codrutas", txtNumero(Index).Text, "N")
            End Select

            If miSQL = "" Then
                MsgBox "No existe el codigo: " & txtNumero(Index).Text, vbExclamation
                
                'Si obligaramos a que existiera el codig
                
            End If
        End If
    End If
    txtDescNumero(Index).Text = miSQL
    
End Sub

Private Sub txtTrab_GotFocus(Index As Integer)
    ConseguirFoco txtTrab(Index), 3
End Sub

Private Sub txtTrab_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpressGnral KeyAscii, 2, True
End Sub

Private Sub txtTrab_LostFocus(Index As Integer)


    txtTrab(Index).Text = Trim(txtTrab(Index).Text)
    Codigo = ""
    miSQL = ""

    If txtTrab(Index).Text <> "" Then
        If IsNumeric(txtTrab(Index).Text) Then
            Codigo = DevuelveDesdeBD(conAri, "nomtraba", "straba", "codtraba", txtTrab(Index).Text, "N")
            If Codigo = "" Then miSQL = "El código no pertence a ningun trabajador"
        Else
            miSQL = "Campo numerico"
        End If
    End If
    
    Me.txtDescTra(Index).Text = Codigo
    If miSQL <> "" Then
        MsgBox miSQL, vbExclamation
        If Index < 2 Then
            txtTrab(Index).Text = ""
            PonerFoco txtTrab(Index)
        End If
    End If
End Sub

Private Sub GenerarEntradaMasivaAccionesComerciales()

    
    Set miRsAux = New ADODB.Recordset
    miSQL = "Select * from scrmtipo where codigo = " & txtAccion(0).Text
    miRsAux.Open miSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRsAux.EOF Then
        MsgBox "No se encuentra la accion: " & txtAccion(0).Text & " " & txtDescAccion(0).Text, vbExclamation
        
    Else
    
        miSQL = "insert into `scrmacciones` (`usuario`,`fechora`,`codclien`,`agente`,`codtraba`,`estado`,"
        miSQL = miSQL & "`tipo`,`medio`,`observaciones`) select '" & DevNombreSQL(vUsu.Login) & "','"
        miSQL = miSQL & Format(txtFecha(0).Text, FormatoFecha) & " " & Format(Now, "hh:mm:ss") & "',"
        miSQL = miSQL & "codclien,codagent," & txtTrab(0).Text & ",0,"
        'tipo, medio observaciones
        miSQL = miSQL & txtAccion(0).Text & "," & DBSet(miRsAux!medio, "T") & "," & DBSet(miRsAux!Observaciones, "T")
        miSQL = miSQL & " FROM sclien where codclien in (" & CadenaDesdeOtroForm & ")"
        ejecutar miSQL, False
    
    End If
    miRsAux.Close
    Set miRsAux = Nothing

End Sub

        
        
