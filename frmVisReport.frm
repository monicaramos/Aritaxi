VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmVisReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      lastProp        =   600
      _cx             =   8281
      _cy             =   5318
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'COmentariio

Public Informe As String
'Public SubInformeConta As String 'SubInforme con conexion a la contabilidad. Conectar a las
                            'tablas de la BDatos correspondiente a la empresa: conta1, conta2, etc.
Public ConSubInforme As Boolean 'Si tiene subinforme ejecta la funcion AbrirSubInforme para enlazar esta a la BD correspondiente


'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean
Public Opcion As Integer
Public ExportarPDF As Boolean
Public EstaImpreso As Boolean


Public NumCopias As Integer ' (RAFA/ALZIRA 31082006) controla el número de copias en un informe de impresion automática

Public SelecImpresora As Boolean


Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim smrpt As CRAXDRT.Report

'Dim Argumentos() As String
Dim PrimeraVez As Boolean


Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
    
    '[Monica]28/09/2012: si me indican seleccionar impresora
    If SelecImpresora Then
         If mrpt.PrinterSetupEx(Me.hwnd) = 0 Then
                EstaImpreso = True
         End If
    Else
         EstaImpreso = True
    End If
    
End Sub

Private Function PuedoCerrar(SegundoIncial As Single) As Boolean
Dim C As Integer
    PuedoCerrar = False
    C = mrpt.PrintingStatus.Progress
    Debug.Print Now & " e:" & C
    If C = 2 Then
        DoEvents
        If Timer - SegundoIncial < 20 Then
            Screen.MousePointer = vbHourglass
            Espera 1
            'If Timer - SegundoIncial > 5 Then
        Else
            PuedoCerrar = True
        End If
    Else
        PuedoCerrar = True
    End If
End Function


Private Sub Form_Activate()
Dim Incio As Single
Dim Fin As Boolean
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
           
            Screen.MousePointer = vbHourglass
            If SoloImprimir Then
                Incio = Timer
                Do
                    Fin = PuedoCerrar(Incio)
                Loop Until Fin
                Set mrpt = Nothing
                Set mapp = Nothing
            End If
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim J As Integer
Dim NomImpre As String

    On Error GoTo Err_Carga
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    Screen.MousePointer = vbHourglass
    Set mapp = CreateObject("CrystalRuntime.Application")
    Set mrpt = mapp.OpenReport(Informe)
       
       
       
    'Conectar a la BD de la Empresa
    For i = 1 To mrpt.Database.Tables.Count
    
        'NUEVO 21 Mayo 2008
        'Puede que alguna tabla este vinculada a ARICONTA
        If LCase(CStr(mrpt.Database.Tables(i).ConnectionProperties.item("DSN"))) = "vconta" Then
            'A conta
            mrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
            If (InStr(1, mrpt.Database.Tables(i).Name, "_") = 0) Then
               mrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & mrpt.Database.Tables(i).Name
            End If
    
        Else
            'A aritaxi
           mrpt.Database.Tables(i).SetLogOnInfo "vAritaxi", vEmpresa.BDAritaxi, vConfig.User, vConfig.password
    '       If InStr(1, Right(mrpt.Database.Tables(i).Name, 2), "_") = 0 Then
'           If InStr(1, mrpt.Database.Tables(i).Name, "_") = 0 Then
                   mrpt.Database.Tables(i).Location = vEmpresa.BDAritaxi & "." & mrpt.Database.Tables(i).Name
'           Else
           If InStr(1, mrpt.Database.Tables(i).Name, "alias") <> 0 Then
                J = InStr(1, mrpt.Database.Tables(i).Name, "_")
                mrpt.Database.Tables(i).Location = vEmpresa.BDAritaxi & "." & Mid(mrpt.Database.Tables(i).Name, 1, J - 1)
           End If
        End If
    Next i

'
'    If SubInformeConta <> "" Then
'        Set smrpt = mrpt.OpenSubreport(SubInformeConta)
'        For i = 1 To smrpt.Database.Tables.Count
'            smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
'            smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
'        Next i
'    End If
    
    'If ConSubInforme Then AbrirSubreport
    AbrirSubreport
    
    PrimeraVez = True
    
    CargaArgumentos
    
    
    mrpt.RecordSelectionFormula = FormulaSeleccion
'    mrpt.RecordSortFields

    If Opcion = 227 Or Opcion = 230 Then
    'Para INforme de Ventas por cliente
        If mrpt.FormulaFields.GetItemByName("pOrden").Text = "{tmpinformes.importe5}" Then
            mrpt.RecordSortFields.item(1).SortDirection = crDescendingOrder
        End If
    End If
    
    
    If ConSubInforme Then
        If Opcion = 228 Or Opcion = 240 Then
             smrpt.RecordSelectionFormula = mrpt.RecordSelectionFormula
        End If
    End If
    
    
    
    
'    If ConSubInforme Then
'        If Opcion = 50 Then
''            If Not (InStr(1, CStr(smrpt.RecordSelectionFormula), "tmpstockfec") > 0) Then
''                smrpt.RecordSelectionFormula = smrpt.RecordSelectionFormula & " and " & mrpt.RecordSelectionFormula
''            End If
'        End If
'    End If
    
    
    'Si es a mail
    If Me.ExportarPDF Then
        Exportar
        Exit Sub
    End If
    
     'lOS MARGENES
'    PonerMargen
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    
    
    EstaImpreso = False
'    mrpt.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
'    If Opcion = 93 Then 'TICKET
'        I = ObtenerTerminal
'        'Establecemos la impresora de ticket
'        NomImpre = NombreImpresoraTicket(I)
'
'        '## PRUEBAS
''        Dim X As Printer
''
'''        oImp = ObtenerImpresora(NomImpre)
''        For Each X In Printers
''           If X.DeviceName = NomImpre Then
''              ' La define como predeterminada del sistema.
''    '          Set Printer = X
''              ' Sale del bucle.
'''              ObtenerImpresora = X
''              Exit For
''           End If
''        Next
''        mrpt.SelectPrinter X.DriverName, X.DeviceName, X.Port
'        '##
'
'        mrpt.SelectPrinter "", NomImpre, ""
'    End If
    
    CRViewer1.ReportSource = mrpt
   
'    If InStr(1, Informe, "Tarj") <> 0 Then
'        If mrpt.PrinterSetupEx(Me.hWnd) = 0 Then
'            mrpt.PrintOut False
'        End If
'    Else
        If SoloImprimir Then
    '        mrpt.PrinterName
    '        Debug.Print mrpt.PrinterName
    
    
            If NumCopias = 0 Then '(RAFA/ALZIRA 31082006) si se ha solicitado número de copias se imprime ese número
                mrpt.PrintOut False
            Else
                mrpt.PrintOut False, NumCopias
            End If
            EstaImpreso = True
        Else
            CRViewer1.ViewReport
        End If
'    End If
    Exit Sub
    
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
    Set smrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim i As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor
    
'    OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
Select Case NumeroParametros
Case 0
    '====Comenta: LAura
    'Solo se vacian los campos de formula que empiezan con "p" ya que estas
    'formulas se corresponden con paso de parametros al Report
    For i = 1 To mrpt.FormulaFields.Count
        If Left(Mid(mrpt.FormulaFields(i).Name, 3), 1) = "p" Then
            mrpt.FormulaFields(i).Text = """"""
        End If
    Next i
    '====
Case 1
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(i).Text = Parametro
        Else
'            mrpt.FormulaFields(I).Text = """"""
        End If
    Next i
    
Case Else
    NumeroParametros = NumeroParametros + 1
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        If DevuelveValor(Parametro) Then
            mrpt.FormulaFields(i).Text = Parametro
        End If
    Next i
'    mrpt.RecordSelectionFormula
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
    Set smrpt = Nothing
    NumCopias = 0 ' (RAFA/ALZIRA 31082006) por si acaso
    SelecImpresora = False
End Sub


Private Function DevuelveValor(ByRef Valor As String) As Boolean
Dim i As Long
Dim J As Long

    Valor = "|" & Valor & "="
    DevuelveValor = False
    i = InStr(1, OtrosParametros, Valor, vbTextCompare)
    If i > 0 Then
        i = i + Len(Valor)
        J = InStr(i, OtrosParametros, "|")
        If J > 0 Then
            Valor = Mid(OtrosParametros, i, J - i)
            If Valor = "" Then
                Valor = " "
            Else
                CompruebaComillas Valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim Aux As String
Dim J As Integer
Dim i As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        Aux = Mid(Valor1, 2, Len(Valor1) - 2)
        i = -1
        Do
            J = i + 2
            i = InStr(J, Aux, """")
            If i > 0 Then
              Aux = Mid(Aux, 1, i - 1) & """" & Mid(Aux, i)
            End If
        Loop Until i = 0
        Aux = """" & Aux & """"
        Valor1 = Aux
    End If
End Sub

Private Sub Exportar()
    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
    mrpt.ExportOptions.DestinationType = crEDTDiskFile
    mrpt.ExportOptions.PDFExportAllPages = True
    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
    mrpt.Export False
    'Si ha generado bien entonces
    CadenaDesdeOtroForm = "OK"
End Sub

Private Sub PonerMargen()
Dim cad As String
Dim i As Integer
    On Error GoTo EPon
    cad = Dir(App.Path & "\*.mrg")
    If cad <> "" Then
        i = InStr(1, cad, ".")
        If i > 0 Then
            cad = Mid(cad, 1, i - 1)
            If IsNumeric(cad) Then
                If Val(cad) > 4000 Then cad = "4000"
                If Val(cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(cad)
                End If
            End If
        End If
    End If
    
    Exit Sub
EPon:
    Err.Clear
End Sub


'======== LAURA ===============================================================
Private Sub AbrirSubreport()
'Para cada subReport que encuentre en el Informe pone las tablas del subReport
'apuntando a la BD correspondiente
Dim crxSection As CRAXDRT.Section
Dim crxObject As Object
Dim crxSubreportObject As CRAXDRT.SubreportObject
Dim i As Byte

    For Each crxSection In mrpt.Sections
        For Each crxObject In crxSection.ReportObjects
             If TypeOf crxObject Is SubreportObject Then
                Set crxSubreportObject = crxObject
                Set smrpt = mrpt.OpenSubreport(crxSubreportObject.SubreportName)
                For i = 1 To smrpt.Database.Tables.Count 'para cada tabla
                    '------ Añade Laura: 09/06/2005
                    If smrpt.Database.Tables(i).ConnectionProperties.item("DSN") = "vAritaxi" Then
                        smrpt.Database.Tables(i).SetLogOnInfo "vAritaxi", vEmpresa.BDAritaxi, vConfig.User, vConfig.password
'                        If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                           smrpt.Database.Tables(i).Location = vEmpresa.BDAritaxi & "." & smrpt.Database.Tables(i).Name
'                        End If
                    ElseIf smrpt.Database.Tables(i).ConnectionProperties.item("DSN") = "vConta" Then
                        smrpt.Database.Tables(i).SetLogOnInfo "vConta", "conta" & vParamAplic.NumeroConta, vParamAplic.UsuarioConta, vParamAplic.PasswordConta
                        If (InStr(1, smrpt.Database.Tables(i).Name, "_") = 0) Then
                           smrpt.Database.Tables(i).Location = "conta" & vParamAplic.NumeroConta & "." & smrpt.Database.Tables(i).Name
                        End If
                    End If
                    '------
                Next i
             End If
        Next crxObject
    Next crxSection
    
    Set crxSubreportObject = Nothing
End Sub

