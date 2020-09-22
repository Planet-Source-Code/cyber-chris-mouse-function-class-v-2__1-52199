VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event PositionChanged()
Public Event SytemClick(ByVal Button As MouseButtonConstants)
Private Const SM_CMOUSEBUTTONS              As Long = 43
Private Const SM_MOUSEWHEELPRESENT          As Long = 75
Private Const SM_MOUSEPRESENT               As Long = 19
Private Const SM_SWAPBUTTON                 As Long = 23
Private Const SPI_SETMOUSESPEED             As Long = 113
Private Const SPI_GETMOUSESPEED             As Long = 112
Private Const GCLNG_CURSOR                  As Long = -12
Private Const VK_RBUTTON                    As Long = &H2
Private Const VK_MBUTTON                    As Long = &H4
Private Const VK_LBUTTON                    As Long = &H1
Private Const MOUSEEVENTF_LEFTDOWN          As Long = &H2
Private Const MOUSEEVENTF_LEFTUP            As Long = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN        As Long = &H20
Private Const MOUSEEVENTF_MIDDLEUP          As Long = &H40
Private Const MOUSEEVENTF_RIGHTDOWN         As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP           As Long = &H10
Private Const OCR_NORMAL                    As Long = 32512
Private Type POINTAPI
    X                                         As Long
    Y                                         As Long
End Type
Private Type CURSORTYPE
    hWnd                                      As Long
    hCursor                                   As Long
    Classname                                 As String
End Type
Private m_Key                               As String
Private m_WatchPosition                     As Boolean
Private m_WatchSystemClicks                 As Boolean
Private m_Position                          As New clsMousePosition
Private m_Clip                              As New clsMouseClip
Private m_Visible                           As Boolean
Private m_Blocked                           As Boolean
Private hNewCursor                          As Long
Private hOldCursor()                        As CURSORTYPE
Private hOldSysCursor                       As Long
Private hNewSysCursor                       As Long
Private AniParent                           As Form
Private AniCursor                           As avbLocationConstants
Private ClassNames                          As String
Public Enum avbLocationConstants
    avbGlobal = 1
    avbLocal = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private avbGlobal, avbLocal
#End If
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
                                              ByVal dx As Long, _
                                              ByVal dy As Long, _
                                              ByVal cButtons As Long, _
                                              ByVal dwExtraInfo As Long)
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
''Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, _
                                                                          ByVal nIndex As Long, _
                                                                          ByVal dwNewLong As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByRef lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Boolean
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Any) As Long
Private Declare Function SetSystemCursor Lib "user32" (ByVal hCur As Long, _
                                                       ByVal id As Long) As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hCur As Long) As Long

''Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Sub Animate(ByRef frmParent As Form, _
                   ByRef Path As String, _
                   Optional ByRef Location As avbLocationConstants = avbLocal)


    If Location = avbLocal Then
        If AniCursor = avbGlobal Then
            Call DeleteAniCursorGlobal
        End If
        Call DeleteAniCursor
        Call CreateAniCursor(frmParent, Path)
     ElseIf Location = avbGlobal Then 'NOT LOCATION...
        If AniCursor = avbLocal Then
            Call DeleteAniCursor
        End If
        Call CreateAniCursorGlobal(Path)
    End If

End Sub

Public Property Get Block() As Boolean

    Block = m_Blocked

End Property

Public Property Let Block(blnValue As Boolean)


    Call BlockInput(IIf(blnValue, 1, 0))
    m_Blocked = blnValue

End Property

Public Property Get Buttons() As Long

    Buttons = GetSystemMetrics(SM_CMOUSEBUTTONS)

End Property

Public Property Get ButtonsSwapped() As Boolean

    ButtonsSwapped = (GetSystemMetrics(SM_SWAPBUTTON) <> 0)

End Property

Public Property Let ButtonsSwapped(blnValue As Boolean)


    Call SwapMouseButton(IIf(blnValue, 1, 0))

End Property

Public Property Get Capture() As Long

    Capture = GetCapture

End Property

Public Property Let Capture(lngHWnd As Long)


    Call SetCapture(lngHWnd)

End Property

Public Sub CaptureRelease()


    Call ReleaseCapture

End Sub

Private Sub Class_Initialize()

    Call modMouse.AddObject(Me, m_Key)
    m_Visible = True

End Sub

Private Sub Class_Terminate()

    If Not m_Visible Then
        Me.Visible = True
    End If
    If m_Blocked Then
        Me.Block = False
    End If
    If AniCursor = avbLocal Then
        Call DeleteAniCursor
     ElseIf AniCursor = avbGlobal Then 'NOT ANICURSOR...
        Call DeleteAniCursorGlobal
    End If
    Call modMouse.RemoveObject(m_Key)
    Set m_Position = Nothing
    Set m_Clip = Nothing

End Sub

Public Sub Click(Optional MouseButton As MouseButtonConstants = vbLeftButton)

    If (MouseButton = vbLeftButton) Then
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
     ElseIf (MouseButton = vbMiddleButton) Then 'NOT (MOUSEBUTTON...
        Call mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_MIDDLEUP, 0&, 0&, 0&, 0&)
     ElseIf (MouseButton = vbRightButton) Then 'NOT (MOUSEBUTTON...
        Call mouse_event(MOUSEEVENTF_RIGHTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_RIGHTUP, 0&, 0&, 0&, 0&)
    End If

End Sub

Public Property Get Clip() As clsMouseClip

    Set Clip = m_Clip

End Property

Private Function CompKey(KCode As Long) As Boolean

  Dim Result As Long

    Result = GetAsyncKeyState(KCode)
    If Result = -32767 Then
        CompKey = True
     Else 'NOT RESULT...
        CompKey = False
    End If

End Function

Private Sub CreateAniCursor(frmParent As Form, _
                            Path As String)

  
  Dim Obj As Object

    If (AniCursor = 0) Then
        If (Dir(Path) <> "") Then
            If (Path <> "") Then
                ReDim hOldCursor(1)
                Set AniParent = frmParent
                hNewCursor = LoadCursorFromFile(Path)
                ClassNames = Chr$(0)
                For Each Obj In AniParent.Controls
                    If HashWnd(Obj) Then
                        If InStr(ClassNames, Chr$(0) & TypeName(Obj) & Chr$(0)) = 0 Then
                            With hOldCursor(UBound(hOldCursor))
                                .hWnd = Obj.hWnd
                                .hCursor = SetClassLong(Obj.hWnd, GCLNG_CURSOR, hNewCursor)
                                .Classname = TypeName(Obj)
                                ClassNames = ClassNames & .Classname & Chr$(0)
                            End With 'HOLDCURSOR(UBOUND(HOLDCURSOR))
                            ReDim Preserve hOldCursor(0 To UBound(hOldCursor) + 1)
                        End If
                    End If
                Next Obj
                With hOldCursor(0)
                    .hWnd = AniParent.hWnd
                    .hCursor = SetClassLong(AniParent.hWnd, GCLNG_CURSOR, hNewCursor)
                    .Classname = TypeName(AniParent)
                    ClassNames = ClassNames & .Classname & Chr$(0)
                End With 'HOLDCURSOR(0)
                AniCursor = avbLocal
            End If
        End If
    End If

End Sub

Private Sub CreateAniCursorGlobal(Path As String)

    If (AniCursor = 0) Then
        If (Dir(Path) <> "") Then
            If (Path <> "") Then
                hOldSysCursor = CopyIcon(GetCursor())
                hNewSysCursor = LoadCursorFromFile(Path)
                Call SetSystemCursor(hNewSysCursor, OCR_NORMAL)
                AniCursor = avbGlobal
            End If
        End If
    End If

End Sub

Public Sub Deactivate()

    Shell "Rundll32.exe mouse,disable"

End Sub

Private Sub DeleteAniCursor()

  Dim Obj As Object
  Dim z   As Long

    If GetUBound(hOldCursor) And (AniCursor = avbLocal) Then
        Call SetClassLong(hOldCursor(0).hWnd, GCLNG_CURSOR, hOldCursor(0).hCursor)
        ClassNames = Replace(ClassNames, Chr$(0) & hOldCursor(0).Classname & Chr$(0), Chr$(0))
        For Each Obj In AniParent.Controls
            If HashWnd(Obj) Then
                If InStr(ClassNames, Chr$(0) & TypeName(Obj) & Chr$(0)) > 0 Then
                    For z = 0 To UBound(hOldCursor) - 1
                        If hOldCursor(z).hWnd = Obj.hWnd Then
                            Call SetClassLong(Obj.hWnd, GCLNG_CURSOR, hOldCursor(z).hCursor)
                            ClassNames = Replace(ClassNames, Chr$(0) & hOldCursor(z).Classname & Chr$(0), Chr$(0))
                            Exit For
                        End If
                    Next z
                End If
            End If
        Next Obj
        Call DestroyCursor(hNewCursor)
        Set AniParent = Nothing
    End If
    AniCursor = 0

End Sub

Private Sub DeleteAniCursorGlobal()

    If AniCursor = avbGlobal Then
        Call SetSystemCursor(hOldSysCursor, OCR_NORMAL)
        Call DestroyCursor(hNewSysCursor)
        Call DestroyCursor(hOldSysCursor)
        AniCursor = 0
    End If

End Sub

Public Property Get DoubleClickTime() As Long

    DoubleClickTime = GetDoubleClickTime

End Property

Public Property Let DoubleClickTime(ByVal ms As Long)

  
  Dim V As Long

    V = ms
    If V < 1 Then
        V = 1
    End If
    If V > 2000 Then
        V = 2000
    End If
    Call SetDoubleClickTime(V)

End Property

Public Property Get Exists() As Long

    Exists = (GetSystemMetrics(SM_MOUSEPRESENT) <> 0)

End Property

Private Function GetUBound(Field() As CURSORTYPE) As Long

    On Error Resume Next
    GetUBound = UBound(Field)
    On Error GoTo 0

End Function

Private Function HashWnd(ByVal Obj As Object) As Boolean


    On Error GoTo Err
    If Obj.hWnd Then
        HashWnd = True
    End If
Err:

End Function

Public Property Get Position() As clsMousePosition

    Set Position = m_Position

End Property

Public Property Get Speed() As Long

  Dim Result As Long

    Call SystemParametersInfo(SPI_GETMOUSESPEED, 0, Result, 0)
    Speed = Result

End Property

Public Property Let Speed(ByVal lngValue As Long)

  
  Dim V As Long

    V = lngValue
    If V < 1 Then
        V = 1
    End If
    If V > 20 Then
        V = 20
    End If
    Call SystemParametersInfo(SPI_SETMOUSESPEED, 0&, ByVal V, 0&)

End Property

Public Property Let TimerEvent(ByVal Dummmy As Boolean)

  
  Dim CurPos   As POINTAPI
  Dim Value    As MouseButtonConstants
  Static First As Boolean

  Static mx    As Long
  Static my    As Long
    If m_WatchPosition Then
        Call GetCursorPos(CurPos)
        If First Then
            If CurPos.X <> mx Or CurPos.Y <> my Then
                RaiseEvent PositionChanged
            End If
        End If
        mx = CurPos.X
        my = CurPos.Y
    End If
    If m_WatchSystemClicks Then
        If CompKey(VK_LBUTTON) Then
            Value = vbLeftButton
        End If
        If CompKey(VK_RBUTTON) Then
            Value = Value Or vbRightButton
        End If
        If CompKey(VK_MBUTTON) Then
            Value = Value Or vbMiddleButton
        End If
        If Value <> 0 Then
            RaiseEvent SytemClick(Value)
        End If
    End If
    First = True

End Property

Public Property Get Visible() As Boolean

    Visible = m_Visible

End Property

Public Property Let Visible(blnValue As Boolean)


    If blnValue <> m_Visible Then
        Call ShowCursor(IIf(blnValue, 1, 0))
        m_Visible = blnValue
    End If

End Property

Public Property Get WatchPosition() As Boolean

    WatchPosition = m_WatchPosition

End Property

Public Property Let WatchPosition(ByVal n_WatchPosition As Boolean)


    m_WatchPosition = n_WatchPosition

End Property

Public Property Get WatchSystemClicks() As Boolean

    WatchSystemClicks = m_WatchSystemClicks

End Property

Public Property Let WatchSystemClicks(ByVal n_WatchSystemClicks As Boolean)


    m_WatchSystemClicks = n_WatchSystemClicks

End Property

Public Property Get WheelExists() As Long

    WheelExists = (GetSystemMetrics(SM_MOUSEWHEELPRESENT) <> 0)

End Property

':)Roja's VB Code Fixer V1.1.92 (07.03.2004 07:57:21) 105 + 436 = 541 Lines Thanks Ulli for inspiration and lots of code.

