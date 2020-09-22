VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMouseClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X                  As Long
    Y                  As Long
End Type
Private Type RECT
    x1                 As Long
    y1                 As Long
    x2                 As Long
    y2                 As Long
End Type
Private m_Left       As Long
Private m_Top        As Long
Private m_Width      As Long
Private m_Height     As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function GetClipCursor Lib "user32" (lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
                                                     lpRect As RECT) As Long

Public Sub Control(ByRef lngHWnd As Long)

  
  Dim R As RECT

    Call GetWindowRect(lngHWnd, R)
    Call ClipCursor(R)

End Sub

Private Sub GetClipRect()

  Dim R As RECT

    Call GetClipCursor(R)
    With R
        m_Left = .x1
        m_Top = .y1
        m_Width = .x2 - .x1
        m_Height = .y2 - .y1
    End With 'R

End Sub

Public Property Get Height() As Long

    Call GetClipRect
    Height = m_Height

End Property

Public Property Let Height(ByRef lngValue As Long)


    m_Height = lngValue
    Call SetClipRect

End Property

Public Property Get Left() As Long

    Call GetClipRect
    Left = m_Left

End Property

Public Property Let Left(ByRef lngValue As Long)


    m_Left = lngValue
    Call SetClipRect

End Property

Public Sub Release()

    Call ClipCursor(ByVal 0&)

End Sub

Private Sub SetClipRect()

  Dim R As RECT

    With R
        .x1 = m_Left
        .y1 = m_Top
        .x2 = m_Width - m_Left
        .y2 = m_Height - m_Top
    End With 'R
    Call ClipCursor(R)

End Sub

Public Property Get Top() As Long

    Call GetClipRect
    Top = m_Top

End Property

Public Property Let Top(ByRef lngValue As Long)


    m_Top = lngValue
    Call SetClipRect

End Property

Public Property Get Width() As Long

    Call GetClipRect
    Width = m_Width

End Property

Public Property Let Width(ByRef lngValue As Long)


    m_Width = lngValue
    Call SetClipRect

End Property

':)Roja's VB Code Fixer V1.1.92 (07.03.2004 07:57:23) 18 + 119 = 137 Lines Thanks Ulli for inspiration and lots of code.

