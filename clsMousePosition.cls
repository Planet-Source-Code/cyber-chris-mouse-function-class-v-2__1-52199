VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMousePosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X            As Long
    Y            As Long
End Type
Private m_x    As Long
Private m_y    As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, _
                                                    ByVal Y As Long) As Long

Private Sub GetPosition()

  Dim P As POINTAPI

    Call GetCursorPos(P)
    m_y = P.Y
    m_x = P.X

End Sub

Public Property Get X() As Long

    Call GetPosition
    X = m_x

End Property

Public Property Let X(lngValue As Long)


    Call SetCursorPos(lngValue, m_y)
    m_x = lngValue

End Property

Public Property Get Y() As Long

    Call GetPosition
    Y = m_y

End Property

Public Property Let Y(lngValue As Long)


    Call SetCursorPos(m_x, lngValue)
    m_y = lngValue

End Property

':)Roja's VB Code Fixer V1.1.92 (07.03.2004 07:57:22) 9 + 47 = 56 Lines Thanks Ulli for inspiration and lots of code.

