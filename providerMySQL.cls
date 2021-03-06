VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "providerMySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents m_pCalendar As CalendarControl
Attribute m_pCalendar.VB_VarHelpID = -1

Dim m_Connection As New ADODB.Connection
Dim m_bOpened As Boolean
Dim m_pHelper As New SQLDataHelper

'======================================================================
Public Sub SetCalendar(ByVal pCalendar As CalendarControl)
    Set m_pCalendar = pCalendar
    
    m_pHelper.SetCalendar pCalendar
End Sub

Public Sub OpenDB(ByVal strConnection As String)
    Debug.Assert Not m_bOpened
    
On Error GoTo Err1
   ' Open DSN connection
    m_Connection.Open m_pHelper.GetDSN(strConnection)
    m_bOpened = True

Err1:
    If Err.Number <> 0 Then
        m_bOpened = False
        MsgBox "Cannot open SQL Server DB" & Err.Description
    End If
End Sub

Public Sub CloseDB()
    If m_bOpened Then
        m_Connection.Close
        m_bOpened = False
    End If
End Sub


Private Sub Class_Initialize()
    m_bOpened = False
End Sub

Private Sub Class_Terminate()
    CloseDB
End Sub


Private Sub m_pCalendar_DoCreateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, NewEventID As Long, bResult As Boolean)
    
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo Err1

    Dim strSQL As String
    strSQL = m_pHelper.MakeCreateEventSQL(pEvent)
    'Debug.Print strSQL
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    '-- update Auto EventID
    Dim rsNewID As ADODB.Recordset
    strSQL = "SELECT MAX(EventID) AS NewEventID FROM CalendarEvents"
    
    Set rsNewID = m_Connection.Execute(strSQL)
    NewEventID = rsNewID("NewEventID")
            
    rsNewID.Close
    
    bResult = True
    Exit Sub

Err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot CreateEvent in DB: " & Err.Description
        Debug.Assert False
    End If
End Sub

Private Sub m_pCalendar_DoCreateRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, NewPatternID As Long, bResult As Boolean)
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo Err1

    Dim strSQL As String
    strSQL = m_pHelper.MakeCreatePatternSQL(pPattern)
    'Debug.Print strSQL
    
    m_Connection.Execute strSQL, , adExecuteNoRecords

    '-- update Auto EventID
    Dim rsNewID As ADODB.Recordset
    strSQL = "SELECT MAX(RecurrencePatternID) AS NewNewPatternID FROM CalendarRecurrencePatterns"
    
    Set rsNewID = m_Connection.Execute(strSQL)
    NewPatternID = rsNewID("NewNewPatternID")
    
    rsNewID.Close
    
    bResult = True
    Exit Sub
Err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
        Debug.Assert False
    End If

End Sub

Private Sub m_pCalendar_DoDeleteEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "DELETE FROM CalendarEvents WHERE EventID = " & pEvent.id
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    bResult = True
End Sub

Private Sub m_pCalendar_DoDeleteRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, bResult As Boolean)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "DELETE FROM CalendarRecurrencePatterns WHERE RecurrencePatternID = " & pPattern.id
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    bResult = True
End Sub

Private Sub m_pCalendar_DoGetUpcomingEvents(ByVal dtFrom As Date, ByVal PeriodMinutes As Long, ByVal pEvents As XtremeCalendarControl.CalendarEvents)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    ' The SQL script below is read all events from DB
    ' RemindersManager will select only reminder which will be Fiered
    ' untill dtFrom + PeriodMinutes. (generally dtFrom is Now)
    '
    ' Of cause you can write this script more optimized.
    '
    
    Dim bOptimizeSQL As Boolean
    bOptimizeSQL = True
    
    Dim strSQL As String
    strSQL = m_pHelper.MakeGetUpcomingEventsSQL(cjCalendarData_MySQL, dtFrom, PeriodMinutes, bOptimizeSQL)
    'Debug.Print strSQL
    
    '----------------------------------------------------------
    Dim rsEvents As ADODB.Recordset
    Set rsEvents = m_Connection.Execute(strSQL)
    
    Dim pEvent As CalendarEvent
        
    While Not rsEvents.EOF
        Set pEvent = m_pHelper.CreateEventFromRS(rsEvents, False)
        If Not pEvent Is Nothing Then
            pEvents.Add pEvent
        End If
        
        rsEvents.MoveNext
    Wend

End Sub

Private Sub m_pCalendar_DoReadEvent(ByVal EventID As Long, pEvent As XtremeCalendarControl.CalendarEvent)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarEvents WHERE EventID = " & EventID
        
    Dim rsEvent As ADODB.Recordset
    Set rsEvent = m_Connection.Execute(strSQL)
    
    Set pEvent = m_pHelper.CreateEventFromRS(rsEvent, False)
End Sub

Private Sub m_pCalendar_DoReadRPattern(ByVal PatternID As Long, pPattern As XtremeCalendarControl.CalendarRecurrencePattern)
    If Not m_bOpened Then
        Exit Sub
    End If
    
    Dim strSQL As String
    strSQL = "SELECT * FROM CalendarRecurrencePatterns "
    strSQL = strSQL + " WHERE RecurrencePatternID = " & PatternID
        
    Dim rsRPattern As ADODB.Recordset
    Set rsRPattern = m_Connection.Execute(strSQL)
    
    Set pPattern = m_pHelper.CreateRPatternFromRS(rsRPattern)
End Sub

Private Sub m_pCalendar_DoRetrieveDayEvents(ByVal dtDay As Date, ByVal pEvents As XtremeCalendarControl.CalendarEvents)

    If Not m_bOpened Then
        Exit Sub
    End If
        
    Dim strSQL As String
    strSQL = m_pHelper.MakeRetrieveDayEventsSQL(dtDay)
    
    Dim rsEvents As ADODB.Recordset
    Set rsEvents = m_Connection.Execute(strSQL)
    
    Dim pEvent As CalendarEvent
        
    While Not rsEvents.EOF
        Set pEvent = m_pHelper.CreateEventFromRS(rsEvents, False)
        If Not pEvent Is Nothing Then
            pEvents.Add pEvent
        End If
        
        rsEvents.MoveNext
    Wend
       
End Sub

Private Sub m_pCalendar_DoUpdateEvent(ByVal pEvent As XtremeCalendarControl.CalendarEvent, bResult As Boolean)
    
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo Err1

    Dim strSQL As String
    
    strSQL = m_pHelper.MakeUpdateEventSQL(pEvent)
    'Debug.Print strSQL
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    bResult = True
    Exit Sub
Err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
        Debug.Assert False
    End If
    
End Sub

Private Sub m_pCalendar_DoUpdateRPattern(ByVal pPattern As XtremeCalendarControl.CalendarRecurrencePattern, bResult As Boolean)
    
    bResult = False
    
    If Not m_bOpened Then
        Exit Sub
    End If

On Error GoTo Err1

    Dim strSQL As String
    strSQL = m_pHelper.MakeUpdatePatternSQL(pPattern)
    'Debug.Print strSQL
    
    m_Connection.Execute strSQL, , adExecuteNoRecords
    
    bResult = True
    Exit Sub
Err1:
    If Err.Number <> 0 Then
        Debug.Print "Cannot UpdateEvent in DB: " & Err.Description
        Debug.Assert False
    End If
    
End Sub


