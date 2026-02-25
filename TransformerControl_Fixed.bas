Option Explicit

'==============================================================================
' CONFIG (edit to match your workbook)
'==============================================================================

Public Const STATE_SHEET As String = "States"
Public Const LOAD_SHEET  As String = "Load Analysis"
Public Const PANEL_SHEET As String = "Panel"

'Live unit state cells (States sheet)
Public Const CELL_D As String = "B2"
Public Const CELL_E As String = "C2"
Public Const CELL_F As String = "D2"

'LIVE lockout output (States sheet) -> 1 when overloaded, else 0
Public Const LOCKOUT_CELL As String = "E2"

'Transformer capacity (Load Analysis)
Public Const TRANSFORMER_CAP_CELL As String = "Q13"

'Editable margins (Load Analysis) - enter 3 or 3% or 0.03 etc.
Public Const OP_MARGIN_CELL As String = "Q17"        'Operational margin used to permit starts
Public Const LOCKOUT_MARGIN_CELL As String = "Q18"   'Lockout threshold margin (positive = allow slight temporary over-cap)

Public Const OP_MARGIN_DEFAULT As Double = 0.03
Public Const LOCKOUT_MARGIN_DEFAULT As Double = 0#

'Timed transitions
Public Const START_SECONDS As Double = 5
Public Const RD_SECONDS    As Double = 5

'Panel remaining capacity output
Public Const PANEL_REMAINING_CAP_CELL As String = "J6"

'==============================================================================
' INTERNALS
'==============================================================================

Private nextTick As Date
Private tickScheduled As Boolean

'Lookup cache
Private lookupCached As Boolean
Private hdrRow As Long
Private colState As Long
Private colStation As Long
Private colD As Long
Private colE As Long
Private colF As Long
Private colKey As Long
Private lastDataRow As Long

'==============================================================================
' STATE HELPERS
'==============================================================================

Private Function WS(ByVal name As String) As Worksheet
    Set WS = ThisWorkbook.Worksheets(name)
End Function

Private Function StateCell(ByVal unitName As String) As Range
    Dim wsS As Worksheet: Set wsS = WS(STATE_SHEET)

    Select Case UCase$(unitName)
        Case "D": Set StateCell = wsS.Range(CELL_D)
        Case "E": Set StateCell = wsS.Range(CELL_E)
        Case "F": Set StateCell = wsS.Range(CELL_F)
        Case Else: Err.Raise vbObjectError + 100, , "Unknown unit " & unitName
    End Select
End Function

Private Function NormalizeStateAbbrev(ByVal rawState As String) As String
    Dim s As String
    s = UCase$(Trim$(rawState))

    Select Case s
        Case "PREHEAT", "PRE-HEAT", "PH"
            NormalizeStateAbbrev = "PH"
        Case "STARTING", "START", "ST"
            NormalizeStateAbbrev = "ST"
        Case "RUNNING", "RUN", "RN"
            NormalizeStateAbbrev = "RN"
        Case "RUNDOWN", "RUN-DOWN", "RD"
            NormalizeStateAbbrev = "RD"
        Case "BASELOAD", "BASE LOAD", "BL"
            NormalizeStateAbbrev = "BL"
        Case Else
            NormalizeStateAbbrev = s
    End Select
End Function

Public Function GetState(ByVal unitName As String) As String
    GetState = NormalizeStateAbbrev(CStr(StateCell(unitName).Value2))
End Function

Public Sub SetState(ByVal unitName As String, ByVal newState As String)
    StateCell(unitName).Value2 = NormalizeStateAbbrev(newState)
End Sub

Private Function IsKnownAbbrev(ByVal s As String) As Boolean
    s = NormalizeStateAbbrev(CStr(s))
    IsKnownAbbrev = (s = "PH" Or s = "ST" Or s = "RN" Or s = "RD" Or s = "BL")
End Function

'==============================================================================
' MARGINS (editable)
' Accepts: 3, 3%, 0.03
'==============================================================================

Private Function ReadPercentCell(ByVal addr As String, ByVal defaultDecimal As Double) As Double
    Dim v As Variant, s As String, n As Double
    v = WS(LOAD_SHEET).Range(addr).Value2

    If IsNumeric(v) Then
        n = CDbl(v)
        If n > 1# Then
            ReadPercentCell = n / 100#
        Else
            ReadPercentCell = n
        End If
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        ReadPercentCell = defaultDecimal
        Exit Function
    End If

    s = Replace(s, "%", "")
    n = Val(s)

    If n > 1# Then
        ReadPercentCell = n / 100#
    Else
        ReadPercentCell = n
    End If
End Function

Public Function OperationalMargin() As Double
    OperationalMargin = ReadPercentCell(OP_MARGIN_CELL, OP_MARGIN_DEFAULT)
End Function

Public Function LockoutMargin() As Double
    LockoutMargin = ReadPercentCell(LOCKOUT_MARGIN_CELL, LOCKOUT_MARGIN_DEFAULT)
End Function

'==============================================================================
' LOOKUP TABLE DETECTION
'==============================================================================

Private Sub EnsureLookupCached()
    If lookupCached Then Exit Sub

    Dim wsL As Worksheet: Set wsL = WS(LOAD_SHEET)
    Dim ur As Range: Set ur = wsL.UsedRange
    If ur Is Nothing Then Err.Raise vbObjectError + 200, , "Load Analysis has no UsedRange."

    Dim r As Long, c As Long
    Dim maxR As Long, maxC As Long
    maxR = ur.Row + ur.Rows.Count - 1
    maxC = ur.Column + ur.Columns.Count - 1

    Dim cellText As String, s1 As String, s2 As String, s3 As String, s4 As String
    Dim found As Boolean: found = False

    For r = ur.Row To maxR
        For c = ur.Column To maxC - 4
            cellText = UCase$(Trim$(CStr(wsL.Cells(r, c).Value2)))
            If cellText = "STATE" Then
                s1 = UCase$(Trim$(CStr(wsL.Cells(r, c + 1).Value2)))
                s2 = UCase$(Trim$(CStr(wsL.Cells(r, c + 2).Value2)))
                s3 = UCase$(Trim$(CStr(wsL.Cells(r, c + 3).Value2)))
                s4 = UCase$(Trim$(CStr(wsL.Cells(r, c + 4).Value2)))

                If (s1 = "STATION" Or s1 = "STAION") And s2 = "D" And s3 = "E" And s4 = "F" Then
                    hdrRow = r
                    colState = c
                    colStation = c + 1
                    colD = c + 2
                    colE = c + 3
                    colF = c + 4
                    found = True
                    Exit For
                End If
            End If
        Next c
        If found Then Exit For
    Next r

    If Not found Then
        Err.Raise vbObjectError + 201, , "Could not find State Loads header row (State|Station|D|E|F)."
    End If

    lastDataRow = hdrRow + 1
    Do While Len(Trim$(CStr(wsL.Cells(lastDataRow, colState).Value2))) > 0 _
       Or Len(Trim$(CStr(wsL.Cells(lastDataRow, colStation).Value2))) > 0
        lastDataRow = lastDataRow + 1
    Loop
    lastDataRow = lastDataRow - 1

    Dim scoreState As Long, scoreStation As Long
    scoreState = 0: scoreStation = 0

    For r = hdrRow + 1 To lastDataRow
        If IsKnownAbbrev(CStr(wsL.Cells(r, colState).Value2)) Then scoreState = scoreState + 1
        If IsKnownAbbrev(CStr(wsL.Cells(r, colStation).Value2)) Then scoreStation = scoreStation + 1
    Next r

    If scoreStation >= scoreState Then
        colKey = colStation
    Else
        colKey = colState
    End If

    lookupCached = True
End Sub

Public Sub ClearLookupCache()
    lookupCached = False
End Sub

'==============================================================================
' LOAD CALCS
'==============================================================================

Public Function TransformerCapacityA() As Double
    TransformerCapacityA = CDbl(Val(CStr(WS(LOAD_SHEET).Range(TRANSFORMER_CAP_CELL).Value2)))
End Function

Private Function FindRowByAbbrev(ByVal stateAbbrev As String) As Long
    EnsureLookupCached

    Dim wsL As Worksheet: Set wsL = WS(LOAD_SHEET)
    Dim r As Long, key As String
    stateAbbrev = NormalizeStateAbbrev(stateAbbrev)

    For r = hdrRow + 1 To lastDataRow
        key = NormalizeStateAbbrev(CStr(wsL.Cells(r, colKey).Value2))
        If key = stateAbbrev Then
            FindRowByAbbrev = r
            Exit Function
        End If
    Next r

    FindRowByAbbrev = 0
End Function

Private Function LoadFromLookup(ByVal stateAbbrev As String, ByVal whichCol As String) As Double
    Dim wsL As Worksheet: Set wsL = WS(LOAD_SHEET)
    Dim r As Long: r = FindRowByAbbrev(stateAbbrev)
    If r = 0 Then
        LoadFromLookup = 0#
        Exit Function
    End If

    Dim v As Variant
    Select Case UCase$(whichCol)
        Case "STATION": v = wsL.Cells(r, colStation).Value2
        Case "STATE":   v = wsL.Cells(r, colState).Value2
        Case "D":       v = wsL.Cells(r, colD).Value2
        Case "E":       v = wsL.Cells(r, colE).Value2
        Case "F":       v = wsL.Cells(r, colF).Value2
        Case Else: Err.Raise vbObjectError + 210, , "Bad whichCol: " & whichCol
    End Select

    If IsNumeric(v) Then
        LoadFromLookup = CDbl(v)
    Else
        LoadFromLookup = CDbl(Val(CStr(v)))
    End If
End Function

Private Function BaseLoadA() As Double
    'BL can be in either State or Station data column depending on workbook layout.
    Dim inStation As Double, inState As Double
    inStation = LoadFromLookup("BL", "STATION")
    inState = LoadFromLookup("BL", "STATE")

    If inStation <> 0# Then
        BaseLoadA = inStation
    Else
        BaseLoadA = inState
    End If
End Function

Public Function TotalLoadA_Current() As Double
    TotalLoadA_Current = BaseLoadA() _
        + LoadFromLookup(GetState("D"), "D") _
        + LoadFromLookup(GetState("E"), "E") _
        + LoadFromLookup(GetState("F"), "F")
End Function

Private Function PredictedLoadIfStart(ByVal unitName As String) As Double
    Dim dSt As String, eSt As String, fSt As String
    dSt = GetState("D"): eSt = GetState("E"): fSt = GetState("F")

    Select Case UCase$(unitName)
        Case "D": dSt = "ST"
        Case "E": eSt = "ST"
        Case "F": fSt = "ST"
    End Select

    PredictedLoadIfStart = BaseLoadA() _
        + LoadFromLookup(dSt, "D") _
        + LoadFromLookup(eSt, "E") _
        + LoadFromLookup(fSt, "F")
End Function

Private Function RemainingCapacityA() As Double
    RemainingCapacityA = TransformerCapacityA() - TotalLoadA_Current()
End Function

'==============================================================================
' OUTPUTS: Remaining capacity to Panel!J6 and Lockout state to States!E2
'==============================================================================

Private Sub UpdateRemainingCapacity()
    WS(PANEL_SHEET).Range(PANEL_REMAINING_CAP_CELL).Value2 = RemainingCapacityA()
End Sub

Private Sub EvaluateLockoutLive()
    'LIVE lockout: write 1/0 to States!E2 based on actual load vs trip level.
    'LockoutMargin is applied ABOVE nominal capacity, e.g. 0.03 => trip at 103% of cap.
    Dim cap As Double, loadNow As Double, tripLevel As Double
    cap = TransformerCapacityA()
    loadNow = TotalLoadA_Current()
    tripLevel = cap * (1 + LockoutMargin())

    If cap <= 0# Then
        WS(STATE_SHEET).Range(LOCKOUT_CELL).Value2 = 1
    Else
        WS(STATE_SHEET).Range(LOCKOUT_CELL).Value2 = IIf(loadNow > tripLevel, 1, 0)
    End If

    UpdateRemainingCapacity
End Sub

'==============================================================================
' START / STOP (simultaneous allowed)
'==============================================================================

Public Sub StartUnit(ByVal unitName As String)
    Dim cap As Double, pred As Double, opLimit As Double

    EvaluateLockoutLive
    If WS(STATE_SHEET).Range(LOCKOUT_CELL).Value2 = 1 Then Exit Sub
    If GetState(unitName) <> "PH" Then Exit Sub

    cap = TransformerCapacityA()
    opLimit = cap * (1 - OperationalMargin())
    pred = PredictedLoadIfStart(unitName)

    If pred > opLimit Then
        EvaluateLockoutLive
        Exit Sub
    End If

    SetState unitName, "ST"
    AddPendingTransition unitName, "RN", START_SECONDS
    StartTickLoop
    EvaluateLockoutLive
End Sub

Public Sub StopUnit(ByVal unitName As String)
    If GetState(unitName) <> "RN" Then
        EvaluateLockoutLive
        Exit Sub
    End If

    SetState unitName, "RD"
    AddPendingTransition unitName, "PH", RD_SECONDS
    StartTickLoop
    EvaluateLockoutLive
End Sub

'==============================================================================
' PENDING TRANSITIONS (Mac-safe, no overflow)
'==============================================================================

Private Sub AddPendingTransition(ByVal unitName As String, ByVal targetState As String, ByVal secondsFromNow As Double)
    Dim dueTime As Date
    dueTime = DateAdd("s", secondsFromNow, Now)

    On Error Resume Next
    ThisWorkbook.Names("Pending_" & UCase$(unitName)).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:="Pending_" & UCase$(unitName), _
        RefersTo:="=""" & Format$(dueTime, "yyyy-mm-dd hh:nn:ss") & "|" & UCase$(targetState) & """", _
        Visible:=False
End Sub

Private Function ParseISODateTime(ByVal s As String, ByRef outDate As Date) As Boolean
    On Error GoTo Bad
    Dim dt() As String, d() As String, t() As String
    dt = Split(Trim$(s), " ")
    If UBound(dt) <> 1 Then GoTo Bad
    d = Split(dt(0), "-")
    t = Split(dt(1), ":")
    If UBound(d) <> 2 Then GoTo Bad
    If UBound(t) <> 2 Then GoTo Bad

    outDate = DateSerial(CInt(d(0)), CInt(d(1)), CInt(d(2))) + TimeSerial(CInt(t(0)), CInt(t(1)), CInt(t(2)))
    ParseISODateTime = True
    Exit Function
Bad:
    ParseISODateTime = False
End Function

Private Sub CheckPending(ByVal unitName As String)
    Dim nm As Name, payload As String, parts() As String
    Dim dueTime As Date, targetState As String

    On Error Resume Next
    Set nm = ThisWorkbook.Names("Pending_" & UCase$(unitName))
    On Error GoTo 0
    If nm Is Nothing Then Exit Sub

    payload = Replace(nm.RefersTo, "=", "")
    payload = Replace(payload, """", "")
    parts = Split(payload, "|")
    If UBound(parts) <> 1 Then Exit Sub

    If Not ParseISODateTime(parts(0), dueTime) Then
        nm.Delete
        Exit Sub
    End If
    targetState = NormalizeStateAbbrev(parts(1))

    If Now >= dueTime Then
        SetState unitName, targetState
        nm.Delete
    End If
End Sub

Private Function NameExists(ByVal nameText As String) As Boolean
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nameText)
    NameExists = Not (nm Is Nothing)
    Set nm = Nothing
    On Error GoTo 0
End Function

Private Function HasAnyPending() As Boolean
    HasAnyPending = NameExists("Pending_D") Or NameExists("Pending_E") Or NameExists("Pending_F")
End Function

Public Sub StartTickLoop()
    On Error Resume Next

    If tickScheduled Then
        Application.OnTime EarliestTime:=nextTick, Procedure:="TickTransitions", Schedule:=False
        tickScheduled = False
    End If

    nextTick = DateAdd("s", 1, Now)
    Application.OnTime EarliestTime:=nextTick, Procedure:="TickTransitions", Schedule:=True
    tickScheduled = True

    On Error GoTo 0
End Sub

Public Sub TickTransitions()
    tickScheduled = False

    CheckPending "D"
    CheckPending "E"
    CheckPending "F"

    EvaluateLockoutLive

    If HasAnyPending() Then StartTickLoop
End Sub

Public Sub Btn_StartD_Click(): StartUnit "D": End Sub
Public Sub Btn_StopD_Click():  StopUnit "D": End Sub
Public Sub Btn_StartE_Click(): StartUnit "E": End Sub
Public Sub Btn_StopE_Click():  StopUnit "E": End Sub
Public Sub Btn_StartF_Click(): StartUnit "F": End Sub
Public Sub Btn_StopF_Click():  StopUnit "F": End Sub

Public Sub Btn_AllStart_Click()
    StartUnit "D"
    StartUnit "E"
    StartUnit "F"
End Sub

Public Sub Btn_AllStop_Click()
    StopUnit "D"
    StopUnit "E"
    StopUnit "F"
End Sub

Public Sub Btn_Refresh_Click()
    ClearLookupCache
    EvaluateLockoutLive
End Sub

Public Sub DebugLoads()
    EnsureLookupCached
    Debug.Print "Key column used for abbreviations: "; IIf(colKey = colStation, "Station/Staion", "State")
    Debug.Print "Cap="; TransformerCapacityA()
    Debug.Print "BL="; BaseLoadA()
    Debug.Print "D(" & GetState("D") & ")="; LoadFromLookup(GetState("D"), "D")
    Debug.Print "E(" & GetState("E") & ")="; LoadFromLookup(GetState("E"), "E")
    Debug.Print "F(" & GetState("F") & ")="; LoadFromLookup(GetState("F"), "F")
    Debug.Print "Total="; TotalLoadA_Current()
    Debug.Print "Remain="; RemainingCapacityA()
    Debug.Print "Lockout="; WS(STATE_SHEET).Range(LOCKOUT_CELL).Value2
End Sub
