Attribute VB_Name = "modMouse"
Option Explicit
Private TimerEnabled     As Boolean
Private colObj           As New Collection
Private hTimer           As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Public Function AddObject(ByRef Obj As clsMouse, _
                          Key As String) As Boolean

  Static Counter As Long

    Call Init
    Counter = Counter + 1
    Key = "x" & CStr(Counter)
    colObj.Add Obj, Key
    AddObject = True

End Function

Private Sub Init()

    If hTimer = 0 Then
        hTimer = SetTimer(0, 0, 20&, AddressOf TimerProc)
        TimerEnabled = True
    End If

End Sub

Private Sub MakeEvents()

  
  Dim Obj As clsMouse
  Dim z   As Long

    For Each Obj In colObj
        Obj.TimerEvent = True
        z = z + 1
    Next Obj
    If z = 0 Then
        Call Terminate
    End If

End Sub

Public Function RemoveObject(ByVal Key As String) As Boolean

  Dim Obj As clsMouse
  Dim z   As Long

    colObj.Remove Key
    For Each Obj In colObj
        z = z + 1
    Next Obj
    If z = 0 Then
        Call Terminate
    End If

End Function

Private Sub Terminate()

    Call KillTimer(0, hTimer)
    TimerEnabled = False
    hTimer = 0

End Sub

Private Sub TimerProc(ByVal lngHWnd As Long, _
                      ByVal Msg As Long, _
                      ByVal idEvent As Long, _
                      ByVal dwTime As Long)

  
  Static Flag As Boolean

    If Not Flag Then
        Flag = True
        Call MakeEvents
        Flag = False
    End If

End Sub

':)Roja's VB Code Fixer V1.1.92 (07.03.2004 07:57:23) 6 + 82 = 88 Lines Thanks Ulli for inspiration and lots of code.

