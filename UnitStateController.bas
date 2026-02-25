Option Explicit

'==============================================================================
' PURPOSE
'   - Simulate Unit D/E/F operational states: PH, ST, RN, RD
'   - Enforce transformer capacity lockout (latched) based on load lookup table
'   - Allow starting/stopping MULTIPLE units at the same time
'   - Show remaining transformer capacity (A) on Panel!J6 (can go negative)
'   - Timed transitions:
'       Start:  PH -> ST -> RN   (after START_SECONDS)
'       Stop:   RN -> RD -> PH   (after RD_SECONDS)
'
' NOTES (Mac VBA)
'   - Fixes Runtime Error 6 (Overflow) by storing pending due time as a
'     normal date-time string ("yyyy-mm-dd hh:nn:ss") instead of numeric serial.
'
' MARGINS (editable in Load Analysis)
'   - Operational margin cell is % (e.g. 3 or 3% or 0.03). Read as decimal.
'   - Lockout margin cell is % (e.g. 0 or 0%).
'==============================================================================

'=========================
' CONFIG
'=========================
Public Const STATE_SHEET As String = "States"
Public Const LOAD_SHEET  As String = "Load Analysis"
Public Const PANEL_SHEET As String = "Panel"

'Live unit state cells (edit if yours differ)
Public Const CELL_D As String = "B2"
Public Const CELL_E As String = "C2"
Public Const CELL_F As String = "D2"

'Latched lockout flag (0/1) stored on States sheet
Public Const LOCKOUT_CELL As String = "E2"

'Transformer capacity cell (you currently use Q13)
Public Const TRANSFORMER_CAP_CELL As String = "Q13"

'Margin input cells (editable on Load Analysis sheet)
' - Put 3 or 3% in OP_MARGIN_CELL
' - Put 0 or 0% in LOCKOUT_MARGIN_CELL
Public Const OP_MARGIN_CELL As String = "Q17"
Public Const LOCKOUT_MARGIN_CELL As String = "Q18"

'Defaults if margin cells are blank/unreadable
Public Const OP_MARGIN_DEFAULT As Double = 0.03
Public Const LOCKOUT_MARGIN_DEFAULT As Double = 0#

'Timed transitions (seconds)
Public Const START_SECONDS As Double = 5    'PH->ST then ST->RN
Public Const RD_SECONDS    As Double = 5    'RN->RD then RD->PH

'If lockout latches, abort any ST immediately to RD (recommended)
Public Const ABORT_START_ON_LOCKOUT As Boolean = True

'Panel output
Public Const PANEL_TOTAL_LOAD_CELL As String = "J5"
Public Const PANEL_REMAINING_CAP_CELL As String = "J6"

'OnTime tick scheduling
Private nextTick As Date
Private tickScheduled As Boolean

'=========================
' Cached lookup table location on Load Analysis
'=========================
Private lookupCached As Boolean
Private hdrRow As Long
Private colState As Long
Private colStation As Long
Private colD As Long
Private colE As Long
Private colF As Long
Private lastDataRow As Long

'==============================================================================
' LIVE STATE HELPERS
'==============================================================================
Private Function StateCell(ByVal unitName As String) As Range
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(STATE_SHEET)
    Select Case UCase$(unitName)
        Case "D": Set StateCell = ws.Range(CELL_D)
        Case "E": Set StateCell = ws.Range(CELL_E)
        Case "F": Set StateCell = ws.Range(CELL_F)
        Case Else: Err.Raise vbObjectError + 100, , "Unknown unit " & unitName
    End Select
End Function

Public Function GetState(ByVal unitName As String) As String
    GetState = UCase$(Trim$(CStr(StateCell(unitName).Value2)))
End Function

Public Sub SetState(ByVal unitName As String, ByVal newState As String)
    StateCell(unitName).Value2 = UCase$(Trim$(newState))
End Sub

'==============================================================================
' LOCKOUT LATCH STORAGE
'==============================================================================
Private Function LockoutCell() As Range
    Set LockoutCell = ThisWorkbook.Worksheets(STATE_SHEET).Range(LOCKOUT_CELL)
End Function

Public Function LockoutActive() As Boolean
    Dim rawText As String
    rawText = Trim$(CStr(LockoutCell.Value2))

    'Avoid numeric narrowing conversions (CLng/CInt) to prevent Overflow.
    LockoutActive = (rawText = "1")
End Function

Public Sub SetLockout(ByVal v As Boolean)
    LockoutCell.Value2 = IIf(v, 1, 0)
End Sub

'==============================================================================
' MARGINS (read from Load Analysis cells; fallback to defaults)
'   Accepts:
'     - 3     => 0.03
'     - 3%    => 0.03
'     - 0.03  => 0.03
'==============================================================================
Private Function ReadPercentCell(ByVal addr As String, ByVal defaultDecimal As Double) As Double
    Dim v As Variant, s As String
    v = ThisWorkbook.Worksheets(LOAD_SHEET).Range(addr).Value2

    On Error GoTo Fallback

    If IsNumeric(v) Then
        Dim pct As Double
        pct = CDbl(v)

        If pct > 1# Then
            ReadPercentCell = pct / 100#
        Else
            ReadPercentCell = pct
        End If
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        ReadPercentCell = defaultDecimal
        Exit Function
    End If

    'Strip % if present
    s = Replace(s, "%", "")
    If Val(s) <> 0 Then
        ReadPercentCell = CDbl(Val(s)) / 100#
    Else
        'Allow 0
        If s = "0" Or s = "0.0" Or s = "0.00" Then
            ReadPercentCell = 0#
        Else
            ReadPercentCell = defaultDecimal
        End If
    End If

    Exit Function

Fallback:
    ReadPercentCell = defaultDecimal
End Function

Public Function OperationalMargin() As Double
    OperationalMargin = ReadPercentCell(OP_MARGIN_CELL, OP_MARGIN_DEFAULT)
End Function

Public Function LockoutMargin() As Double
    LockoutMargin = ReadPercentCell(LOCKOUT_MARGIN_CELL, LOCKOUT_MARGIN_DEFAULT)
End Function

'==============================================================================
' LOOKUP TABLE AUTO-DETECTION
'   Finds the header row containing:
'       State | Station(or Staion) | D | E | F
'==============================================================================
Private Sub EnsureLookupCached()
    If lookupCached Then Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOAD_SHEET)
    Dim ur As Range

    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Err.Raise vbObjectError + 200, , "Load Analysis sheet has no UsedRange."

    Dim r As Long, c As Long
    Dim maxR As Long, maxC As Long
    maxR = ur.Row + ur.Rows.Count - 1
    maxC = ur.Column + ur.Columns.Count - 1

    Dim cellText As String, s1 As String, s2 As String, s3 As String, s4 As String
    Dim found As Boolean: found = False

    For r = ur.Row To maxR
        For c = ur.Column To maxC - 4
            cellText = UCase$(Trim$(CStr(ws.Cells(r, c).Value2)))
            If cellText = "STATE" Then
                s1 = UCase$(Trim$(CStr(ws.Cells(r, c + 1).Value2))) 'Station/Staion
                s2 = UCase$(Trim$(CStr(ws.Cells(r, c + 2).Value2))) 'D
                s3 = UCase$(Trim$(CStr(ws.Cells(r, c + 3).Value2))) 'E
                s4 = UCase$(Trim$(CStr(ws.Cells(r, c + 4).Value2))) 'F

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
        Err.Raise vbObjectError + 201, , "Could not find State Loads header row (State|Station|D|E|F) on 'Load Analysis'."
    End If

    'Find last data row by scanning down the State column until blank
    lastDataRow = hdrRow + 1
    Do While Len(Trim$(CStr(ws.Cells(lastDataRow, colState).Value2))) > 0
        lastDataRow = lastDataRow + 1
    Loop
    lastDataRow = lastDataRow - 1

    lookupCached = True
End Sub

'==============================================================================
' LOAD CALCS
'==============================================================================
Public Function TransformerCapacityA() As Double
    TransformerCapacityA = CDbl(Val(ThisWorkbook.Worksheets(LOAD_SHEET).Range(TRANSFORMER_CAP_CELL).Value2))
End Function

Private Function LoadFromLookup(ByVal stateAbbrev As String, ByVal whichCol As String) As Double
    EnsureLookupCached

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOAD_SHEET)
    Dim targetCol As Long

    Select Case UCase$(whichCol)
        Case "STATION": targetCol = colStation
        Case "D": targetCol = colD
        Case "E": targetCol = colE
        Case "F": targetCol = colF
        Case Else: Err.Raise vbObjectError + 210, , "Bad whichCol: " & whichCol
    End Select

    Dim r As Long, key As String
    stateAbbrev = UCase$(Trim$(stateAbbrev))

    For r = hdrRow + 1 To lastDataRow
        key = UCase$(Trim$(CStr(ws.Cells(r, colState).Value2)))
        If key = stateAbbrev Then
            Dim v As Variant
            v = ws.Cells(r, targetCol).Value2
            If IsNumeric(v) Then
                LoadFromLookup = CDbl(v)
            Else
                LoadFromLookup = 0#
            End If
            Exit Function
        End If
    Next r

    LoadFromLookup = 0#
End Function

Private Function BaseLoadA() As Double
    BaseLoadA = LoadFromLookup("BL", "STATION")
End Function

Public Function TotalLoadA_Current() As Double
    TotalLoadA_Current = BaseLoadA() _
        + LoadFromLookup(GetState("D"), "D") _
        + LoadFromLookup(GetState("E"), "E") _
        + LoadFromLookup(GetState("F"), "F")
End Function

Public Function RemainingCapacityA_Current() As Double
    RemainingCapacityA_Current = TransformerCapacityA() - TotalLoadA_Current()
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

'==============================================================================
' PANEL OUTPUT
'==============================================================================
Public Sub UpdateRemainingCapacity()
    Dim loadNow As Double, remainA As Double

    loadNow = TotalLoadA_Current()
    remainA = TransformerCapacityA() - loadNow  'can go negative

    With ThisWorkbook.Worksheets(PANEL_SHEET)
        .Range(PANEL_TOTAL_LOAD_CELL).Value2 = loadNow
        .Range(PANEL_REMAINING_CAP_CELL).Value2 = remainA
    End With
End Sub

'==============================================================================
' LOCKOUT (LATCHING)
'==============================================================================
Private Sub EvaluateAndLatchLockout()
    Dim remainA As Double

    remainA = RemainingCapacityA_Current()

    If (Not LockoutActive()) And (remainA < 0#) Then
        SetLockout True

        If ABORT_START_ON_LOCKOUT Then
            If GetState("D") = "ST" Then SetState "D", "RD"
            If GetState("E") = "ST" Then SetState "E", "RD"
            If GetState("F") = "ST" Then SetState "F", "RD"
        End If
    End If

    UpdateRemainingCapacity
End Sub

Public Sub ResetLockout()
    If GetState("D") = "ST" Or GetState("E") = "ST" Or GetState("F") = "ST" Then Exit Sub

    If RemainingCapacityA_Current() >= 0# Then
        SetLockout False
    End If

    UpdateRemainingCapacity
End Sub

'==============================================================================
' START / STOP COMMANDS
'   Simultaneous starts/stops allowed.
'==============================================================================
Public Sub StartUnit(ByVal unitName As String)
    Dim cap As Double, pred As Double, opLimit As Double

    EvaluateAndLatchLockout
    If LockoutActive() Then Exit Sub
    If GetState(unitName) <> "PH" Then Exit Sub

    cap = TransformerCapacityA()
    opLimit = cap * (1 - OperationalMargin())
    pred = PredictedLoadIfStart(unitName)

    If pred > opLimit Then
        UpdateRemainingCapacity
        Exit Sub
    End If

    SetState unitName, "ST"
    AddPendingTransition unitName, "RN", START_SECONDS
    StartTickLoop
    UpdateRemainingCapacity
End Sub

Public Sub StopUnit(ByVal unitName As String)
    If GetState(unitName) <> "RN" Then
        UpdateRemainingCapacity
        Exit Sub
    End If

    SetState unitName, "RD"
    AddPendingTransition unitName, "PH", RD_SECONDS
    StartTickLoop
    UpdateRemainingCapacity
End Sub

'==============================================================================
' PENDING TRANSITIONS (OnTime) -- Mac-safe (no numeric serial conversions)
'==============================================================================
Private Sub AddPendingTransition(ByVal unitName As String, ByVal targetState As String, ByVal secondsFromNow As Double)
    Dim dueTime As Date
    dueTime = DateAdd("s", secondsFromNow, Now)

    On Error Resume Next
    ThisWorkbook.Names("Pending_" & UCase$(unitName)).Delete
    On Error GoTo 0

    'Store: "yyyy-mm-dd hh:nn:ss|RN"
    ThisWorkbook.Names.Add _
        Name:="Pending_" & UCase$(unitName), _
        RefersTo:="=""" & Format$(dueTime, "yyyy-mm-dd hh:nn:ss") & "|" & UCase$(targetState) & """", _
        Visible:=False
End Sub

Private Function ParseIsoDateTime(ByVal dtText As String, ByRef parsedDate As Date) As Boolean
    Dim datePart As String, timePart As String
    Dim dBits() As String, tBits() As String
    Dim yy As Long, mm As Long, dd As Long
    Dim hh As Long, nn As Long, ss As Long
    Dim sepPos As Long

    On Error GoTo ParseFail

    dtText = Trim$(dtText)
    If Len(dtText) = 0 Then Exit Function

    sepPos = InStr(1, dtText, " ")
    If sepPos <= 1 Then Exit Function

    datePart = Left$(dtText, sepPos - 1)
    timePart = Mid$(dtText, sepPos + 1)

    dBits = Split(datePart, "-")
    tBits = Split(timePart, ":")

    If UBound(dBits) <> 2 Or UBound(tBits) <> 2 Then Exit Function

    If Not TryParseLong(dBits(0), yy) Then Exit Function
    If Not TryParseLong(dBits(1), mm) Then Exit Function
    If Not TryParseLong(dBits(2), dd) Then Exit Function
    If Not TryParseLong(tBits(0), hh) Then Exit Function
    If Not TryParseLong(tBits(1), nn) Then Exit Function
    If Not TryParseLong(tBits(2), ss) Then Exit Function

    If yy < 1900 Or yy > 9999 Then Exit Function
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function
    If hh < 0 Or hh > 23 Then Exit Function
    If nn < 0 Or nn > 59 Then Exit Function
    If ss < 0 Or ss > 59 Then Exit Function

    parsedDate = DateSerial(yy, mm, dd) + TimeSerial(hh, nn, ss)
    ParseIsoDateTime = True
    Exit Function

ParseFail:
    ParseIsoDateTime = False
End Function

Private Function TryParseLong(ByVal textValue As String, ByRef parsed As Long) As Boolean
    Dim clean As String

    On Error GoTo ParseFail

    clean = Trim$(textValue)
    If Len(clean) = 0 Then Exit Function

    If Not IsNumeric(clean) Then Exit Function

    parsed = CLng(Val(clean))
    TryParseLong = True
    Exit Function

ParseFail:
    TryParseLong = False
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

    If UBound(parts) <> 1 Then
        nm.Delete
        Exit Sub
    End If

    If Not ParseIsoDateTime(parts(0), dueTime) Then
        nm.Delete
        Exit Sub
    End If

    targetState = UCase$(Trim$(parts(1)))
    If targetState <> "PH" And targetState <> "ST" And targetState <> "RN" And targetState <> "RD" Then
        nm.Delete
        Exit Sub
    End If

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

    'Cancel previous tick if scheduled
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

    EvaluateAndLatchLockout
    UpdateRemainingCapacity

    If HasAnyPending() Then StartTickLoop
End Sub

'==============================================================================
' BUTTON MACROS (assign Excel shapes/buttons to these)
'==============================================================================
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

Public Sub Btn_ResetLockout_Click()
    ResetLockout
End Sub

Public Sub Btn_UpdateCapacity_Click()
    UpdateRemainingCapacity
End Sub

'==============================================================================
' DIAGNOSTICS
'==============================================================================
Public Sub DebugLoads()
    Dim cap As Double, baseA As Double, dA As Double, eA As Double, fA As Double, tot As Double
    cap = TransformerCapacityA()
    baseA = BaseLoadA()
    dA = LoadFromLookup(GetState("D"), "D")
    eA = LoadFromLookup(GetState("E"), "E")
    fA = LoadFromLookup(GetState("F"), "F")
    tot = baseA + dA + eA + fA

    Debug.Print "=== DebugLoads ==="
    Debug.Print "Capacity(" & TRANSFORMER_CAP_CELL & ") = "; cap
    Debug.Print "OpMargin(" & OP_MARGIN_CELL & ")      = "; OperationalMargin()
    Debug.Print "Lockout(" & LOCKOUT_MARGIN_CELL & ")  = "; LockoutMargin()
    Debug.Print "Base(BL)                              = "; baseA
    Debug.Print "D load                                = "; dA; " state="; GetState("D")
    Debug.Print "E load                                = "; eA; " state="; GetState("E")
    Debug.Print "F load                                = "; fA; " state="; GetState("F")
    Debug.Print "TOTAL                                 = "; tot
    Debug.Print "Remaining(A)                          = "; cap - tot
End Sub

Public Sub ClearLookupCache()
    lookupCached = False
End Sub
