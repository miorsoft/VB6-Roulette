VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Tick Class Module by Vesa Piittinen < http://merri.net/ >
' ---------------------------------------------------------
' Makes it possible to run portions of code X times a second!

Option Explicit

' default setting for late indicator
Private Const DEFAULT_LATETICKS As Currency = 20

' for better DoEvents handling
Private Const QS_KEY = &H1&
Private Const QS_MOUSEMOVE = &H2&
Private Const QS_MOUSEBUTTON = &H4&
Private Const QS_POSTMESSAGE = &H8&
Private Const QS_TIMER = &H10&
Private Const QS_PAINT = &H20&
Private Const QS_SENDMESSAGE = &H40&
Private Const QS_HOTKEY = &H80&
Private Const QS_ALLINPUT = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)
Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)

' tick data
Private Type TICKS
    ID            As Long          ' identifier ID
    Count         As Long          ' number of ticks done
    Ending        As Currency      ' the next ending tick
    FREQ          As Currency      ' tick frequency
    NoSkip        As Boolean       ' should always (atleast try) to stay on time?
End Type

' internal variables
Dim m_curFreq     As Currency
Dim m_curLateTick As Currency
Dim m_curLateTicks As Currency
Dim m_dblFreqToMS As Double

Private tick()    As TICKS

' for better DoEvents handling
Private Declare Function GetQueueStatus Lib "user32" (ByVal qsFlags As Long) As Long
' for timing
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
' to prevent 100% processor usage
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' add a new timed event
Public Function Add(ByVal FramesPerSecond As Double, Optional ByVal ID As Long = 0, Optional ByVal NoFrameSkip As Boolean = False) As Long
    Dim blnNoArrayInit As Boolean, lngNewIndex As Long
    ' check for invalid input values
    If ID < 0 Then Add = -1: Exit Function
    If FramesPerSecond <= 0 Then Add = -1: Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' if array is initialized, we can get UBound
    If Not blnNoArrayInit Then lngNewIndex = UBound(tick) + 1
    ' add new item
    ReDim Preserve tick(lngNewIndex)
    ' set settings
    With tick(lngNewIndex)
        .ID = ID
        .FREQ = CCur(CDbl(m_curFreq) / FramesPerSecond)
    End With
    ' return the new index
    Add = lngNewIndex
End Function
' how many times a timed event has occurred?
Public Function Count(ByVal index As Long) As Long
    Dim blnNoArrayInit As Boolean
    ' invalid index?
    If index < 0 Then Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, nothing to remove
    If blnNoArrayInit Then Exit Function
    ' out of upper bound?
    If index > UBound(tick) Then Exit Function
    ' finally... return the count
    Count = tick(index).Count
End Function
' how many times timed events of certain identifier have occurred?
Public Function CountByID(ByVal ID As Long) As Long
    Dim blnNoArrayInit As Boolean, lngA As Long, lngCount As Long
    ' invalid ID?
    If ID < 0 Then Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, nothing to remove
    If blnNoArrayInit Then Exit Function
    ' count totals by ID
    For lngA = 0 To UBound(tick)
        If tick(lngA).ID = ID Then lngCount = lngCount + tick(lngA).Count
    Next lngA
    ' finally... return the count
    CountByID = lngCount
End Function
' late indicator: changing this value sets how eagerly event is considered to be late
Public Property Get LateIndicator() As Currency
    LateIndicator = m_curLateTicks
End Property
' smaller value = less eager, bigger value = more eager
' bigger value means an event is moved easier further into the future
Public Property Let LateIndicator(ByVal NewValue As Currency)
    If NewValue <= 1 Then Exit Property
    m_curLateTicks = NewValue
    m_curLateTick = -(m_curFreq / m_curLateTicks)
End Property
' remove a timed event
Public Function Remove(ByVal index As Long) As Boolean
    Dim blnNoArrayInit As Boolean, lngA As Long, lngTicks As Long
    If index < 0 Then Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, nothing to remove
    If blnNoArrayInit Then Exit Function
    ' out of upper bound?
    lngTicks = UBound(tick) - 1
    If (index - 1) > lngTicks Then Exit Function
    ' finally, remove
    If lngTicks >= 0 Then
        ' remove the current index by overwriting
        For lngA = index To lngTicks
            tick(index) = tick(index + 1)
        Next lngA
        ' remove last item
        ReDim Preserve tick(lngTicks)
    Else
        ' remove whole array
        Erase tick
    End If
    ' success
    Remove = True
End Function
' remove timed events of certain identifier
Public Function RemoveByID(ByVal ID As Long) As Long
    Dim blnNoArrayInit As Boolean, lngA As Long, lngB As Long, lngCount As Long
    If ID < 0 Then Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, nothing to remove
    If blnNoArrayInit Then Exit Function
    ' loop through all items
    For lngA = UBound(tick) To 0 Step -1
        ' remove ones with a matching ID
        If tick(lngA).ID = ID Then
            If UBound(tick) > 0 Then
                ' remove current item
                For lngB = lngA To UBound(tick) - 1
                    tick(lngB) = tick(lngB + 1)
                Next lngB
                ' remove the last item of the array
                ReDim Preserve tick(UBound(tick) - 1)
                ' increase counter
                lngCount = lngCount + 1
            Else
                ' remove the last item in array
                Erase tick
                ' increase counter
                lngCount = lngCount + 1
                Exit For
            End If
        End If
    Next lngA
    ' return number of removed items
    RemoveByID = lngCount
End Function
' reset event amount counter
Public Function ResetCount(ByVal index As Long) As Boolean
    Dim blnNoArrayInit As Boolean
    ' invalid index?
    If index < 0 Then Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, nothing to remove
    If blnNoArrayInit Then Exit Function
    ' out of upper bound?
    If index > UBound(tick) Then Exit Function
    ' finally... reset the count
    tick(index).Count = 0
    ResetCount = True
End Function
' reset event amount counter of certain identifier
Public Function ResetCountByID(ByVal ID As Long) As Long
    Dim blnNoArrayInit As Boolean, lngA As Long, lngCount As Long
    ' invalid ID?
    If ID < 0 Then Exit Function
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, nothing to remove
    If blnNoArrayInit Then Exit Function
    ' reset counts by ID
    For lngA = 0 To UBound(tick)
        If tick(lngA).ID = ID Then
            tick(lngA).Count = 0
            lngCount = lngCount + 1
        End If
    Next lngA
    ' finally... return the amount of items we reseted
    ResetCountByID = lngCount
End Function
' this can be used to make ticks occur based on same time
' if not used, all ticks start off at slightly different times
Public Function Start() As Boolean
    Dim blnNoArrayInit As Boolean, curTick As Currency, lngA As Long
    ' check if array is initialized
    blnNoArrayInit = (Not tick) = -1
    ' skip IDE error...
    On Error Resume Next: Debug.Assert 0.1: On Error GoTo 0
    ' no array, we can't start
    If blnNoArrayInit Then Exit Function
    ' get current tick
    QueryPerformanceCounter curTick
    ' set starting time for all
    For lngA = 0 To UBound(tick)
        tick(lngA).Ending = curTick
    Next lngA
    ' success!
    Start = True
End Function
' this returns the INDEX of the next tick
Public Function WaitForNext() As Long
    Dim curDifference As Currency, curTick As Currency, lngA As Long, lngIndex As Long
    ' WARNING! WE HAVE NO ERROR DETECTION! THIS FUNCTION ASSUMES THERE ARE TICKS!
    ' process DoEvents only if needed to
    If GetQueueStatus(QS_SENDMESSAGE Or QS_ALLEVENTS) <> 0 Then DoEvents
    ' figure out the next tick
    For lngA = 1 To UBound(tick)
        If tick(lngA).Ending < tick(lngIndex).Ending Then lngIndex = lngA
    Next lngA
    ' increase counter
    tick(lngIndex).Count = tick(lngIndex).Count + 1
    ' get current tick
    QueryPerformanceCounter curTick
    ' because of the falling behind prevention,
    ' we need to set this if it is not initialized
    If tick(lngIndex).Ending = 0 Then tick(lngIndex).Ending = curTick
    ' then wait for the tick
    curDifference = tick(lngIndex).Ending - curTick
    ' check if we are late or in advance; or perfectly on time
    If curDifference >= 0 Then
        ' we might be early, so we have to wait a bit
        lngA = CLng(CDbl(curDifference) * m_dblFreqToMS)
        ' wait for a few milliseconds if necessary
        If lngA > 0 Then Sleep lngA
        ' set the next time
        tick(lngIndex).Ending = tick(lngIndex).Ending + tick(lngIndex).FREQ
    Else
        If Not tick(lngIndex).NoSkip Then
            ' we are late, but by how much?
            If curDifference > m_curLateTick Then
                ' not too badly, we can process the next tick right on time
                tick(lngIndex).Ending = tick(lngIndex).Ending + tick(lngIndex).FREQ
            Else
                ' we are so badly late in processing that we must SKIP PROCESSING
                ' if we wouldn't do this, then slow computers would just be all too slow
                tick(lngIndex).Ending = curTick + tick(lngIndex).FREQ
            End If
        Else
            ' we are late, but this tick never skips
            tick(lngIndex).Ending = tick(lngIndex).Ending + tick(lngIndex).FREQ
        End If
    End If
    ' return index
    WaitForNext = lngIndex
End Function
' this returns the ID of the next tick
Public Function WaitForNextID() As Long
    Dim curDifference As Currency, curTick As Currency, lngA As Long, lngIndex As Long
    ' WARNING! WE HAVE NO ERROR DETECTION! THIS FUNCTION ASSUMES THERE ARE TICKS!
    ' process DoEvents only if needed to
    If GetQueueStatus(QS_SENDMESSAGE Or QS_ALLEVENTS) <> 0 Then DoEvents
    ' figure out the next tick
    For lngA = 1 To UBound(tick)
        If tick(lngA).Ending < tick(lngIndex).Ending Then lngIndex = lngA
    Next lngA
    ' increase counter
    tick(lngIndex).Count = tick(lngIndex).Count + 1
    ' get current tick
    QueryPerformanceCounter curTick
    ' because of the falling behind prevention,
    ' we need to set this if it is not initialized
    If tick(lngIndex).Ending = 0 Then tick(lngIndex).Ending = curTick
    ' then wait for the tick
    curDifference = tick(lngIndex).Ending - curTick
    ' check if we are late or in advance; or perfectly on time
    If curDifference >= 0 Then
        ' we might be early, so we have to wait a bit
        lngA = CLng(CDbl(curDifference) * m_dblFreqToMS)
        ' wait for a few milliseconds if necessary
        If lngA > 0 Then Sleep lngA
        ' set the next time
        tick(lngIndex).Ending = tick(lngIndex).Ending + tick(lngIndex).FREQ
    Else
        If Not tick(lngIndex).NoSkip Then
            ' we are late, but by how much?
            If curDifference > m_curLateTick Then
                ' not too badly, we can process the next tick right on time
                tick(lngIndex).Ending = tick(lngIndex).Ending + tick(lngIndex).FREQ
            Else
                ' we are so badly late in processing that we must SKIP PROCESSING
                ' if we wouldn't do this, then slow computers would just be all too slow
                tick(lngIndex).Ending = curTick + tick(lngIndex).FREQ
            End If
        Else
            ' we are late, but this tick never skips
            tick(lngIndex).Ending = tick(lngIndex).Ending + tick(lngIndex).FREQ
        End If
    End If
    ' return id
    WaitForNextID = tick(lngIndex).ID
End Function
Private Sub Class_Initialize()
    ' get frequency (= length of a second)
    QueryPerformanceFrequency m_curFreq
    ' for converting frequency to MS
    m_dblFreqToMS = 1000 / CDbl(m_curFreq)
    ' set the default late indicator value
    m_curLateTicks = DEFAULT_LATETICKS
    m_curLateTick = -(m_curFreq / m_curLateTicks)
End Sub
Private Sub Class_Terminate()
    ' clear memory
    Erase tick
End Sub
