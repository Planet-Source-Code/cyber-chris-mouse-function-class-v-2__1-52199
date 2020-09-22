Attribute VB_Name = "modMousewheel"
Option Explicit
Private Const WM_MOUSEWHEEL      As Long = &H20A
Private MouseWheelUp             As Boolean
Public Const GWL_WNDPROC         As Long = (-4)
Public OldProc1                  As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                           ByVal nIndex As Long, _
                                                                           ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                           ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hWnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long

Public Function TWndProc1(ByVal lngHWnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long


    If wMsg = WM_MOUSEWHEEL Then
        If wParam > 0 Then
            MouseWheelUp = True
         Else
            MouseWheelUp = False
        End If
        Select Case MouseWheelUp
         Case True
            frmMouseTest.Label1.Caption = "The value of the mouse wheel has increased"
            If frmMouseTest.pbar.Value < 100 Then frmMouseTest.pbar.Value = frmMouseTest.pbar.Value + 1
         Case False
            With frmMouseTest
                .Label1.Caption = "The value of the mouse wheel has decreased"
                If .pbar.Value > 0 Then
                    .pbar.Value = .pbar.Value - 1
                End If
            End With
        End Select
    End If
    TWndProc1 = CallWindowProc(OldProc1, lngHWnd, wMsg, wParam, lParam)

End Function

