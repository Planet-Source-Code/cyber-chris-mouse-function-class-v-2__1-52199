VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMouseTest 
   Caption         =   "Mouse Function"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form2"
   ScaleHeight     =   7500
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame6 
      Caption         =   "Mouse Wheel"
      Height          =   1215
      Left            =   120
      TabIndex        =   27
      Top             =   6240
      Width           =   7095
      Begin MSComctlLib.ProgressBar pbar 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   6855
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Control"
      Height          =   825
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   7035
      Begin VB.CommandButton cmdEnd 
         Caption         =   "End"
         Height          =   390
         Left            =   5895
         TabIndex        =   26
         Top             =   255
         Width           =   990
      End
      Begin VB.CommandButton cmdSetClick 
         Caption         =   "Set && Click"
         Height          =   390
         Left            =   135
         TabIndex        =   23
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   390
         Left            =   1155
         TabIndex        =   22
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdSwap 
         Caption         =   "Swap"
         Height          =   390
         Left            =   2175
         TabIndex        =   21
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdDeactivate 
         Caption         =   "Deactivate"
         Height          =   390
         Left            =   4215
         TabIndex        =   20
         Top             =   270
         Width           =   1005
      End
      Begin VB.CommandButton cmdBlock 
         Caption         =   "Block"
         Height          =   390
         Left            =   3195
         TabIndex        =   19
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Timer tmrBlock 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4560
      Top             =   1170
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4095
      Top             =   1170
   End
   Begin VB.Frame Frame3 
      Caption         =   "Info"
      Enabled         =   0   'False
      Height          =   1920
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   1800
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   1005
         Width           =   1185
      End
      Begin VB.CheckBox chkSwap 
         Caption         =   "Buttons swapped"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   780
         Width           =   1590
      End
      Begin VB.CheckBox chkWheel 
         Caption         =   "Mousewheel"
         Height          =   300
         Left            =   135
         TabIndex        =   13
         Top             =   510
         Width           =   1245
      End
      Begin VB.CheckBox chkExists 
         Caption         =   "Mouse"
         Height          =   360
         Left            =   135
         TabIndex        =   12
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblButtons 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Label4"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   1350
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Speed"
      Height          =   990
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   7050
      Begin VB.CommandButton cmdSpeedReset 
         Caption         =   "Reset"
         Height          =   585
         Left            =   6240
         TabIndex        =   10
         Top             =   240
         Width           =   780
      End
      Begin VB.HScrollBar scrSpeed 
         Height          =   255
         Left            =   1110
         Max             =   20
         Min             =   1
         TabIndex        =   4
         Top             =   255
         Value           =   1
         Width           =   4110
      End
      Begin VB.HScrollBar scrDblClick 
         Height          =   255
         Left            =   1110
         Max             =   2000
         Min             =   1
         TabIndex        =   3
         Top             =   570
         Value           =   1
         Width           =   4110
      End
      Begin VB.Label lblSpeed 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Left            =   5400
         TabIndex        =   9
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblDblClick 
         BorderStyle     =   1  'Fest Einfach
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   570
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "DoubleClick"
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   585
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "Mouse"
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Animate Cursor"
      Height          =   990
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   1800
      Begin VB.CommandButton cmdAniLocal 
         Caption         =   "Set Form"
         Height          =   300
         Left            =   120
         MousePointer    =   99  'Benutzerdefiniert
         TabIndex        =   7
         Top             =   255
         Width           =   1545
      End
      Begin VB.CommandButton cmdAniGlobal 
         Caption         =   "Set System"
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   1545
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Clip"
      Height          =   1800
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   7050
      Begin VB.PictureBox picClip 
         BackColor       =   &H00FFFFFF&
         Height          =   1380
         Left            =   135
         ScaleHeight     =   1320
         ScaleWidth      =   6720
         TabIndex        =   24
         Top             =   270
         Width           =   6780
         Begin VB.CommandButton cmdClip 
            Caption         =   "Clip Cursor"
            Height          =   390
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1170
         End
      End
   End
End
Attribute VB_Name = "frmMouseTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Advanced Mouse Function Classes
'   (c) Copyright by Cyber Chris
'   cyber_chris235@gmx.net
Option Explicit
Private WithEvents Mouse      As clsMouse
Attribute Mouse.VB_VarHelpID = -1
Private Flag1                 As Boolean
Private Flag2                 As Boolean
Private memSpeed              As Long
Private memDblClick           As Long

Private Sub cmdAniGlobal_Click()

    Flag2 = Not Flag2
    If Flag2 Then
        Mouse.Animate Me, App.Path & "\hourglas.ani", avbGlobal
        cmdAniGlobal.Caption = "Reset System"
     Else 'FLAG2 = FALSE/0
        Mouse.Animate Me, ""
        cmdAniGlobal.Caption = "Set System"
    End If
    Flag1 = False
    cmdAniLocal.Caption = "Set Form"

End Sub

Private Sub cmdAniLocal_Click()

    Flag1 = Not Flag1
    If Flag1 Then
        Mouse.Animate Me, App.Path & "\hourglas.ani", avbLocal
        cmdAniLocal.Caption = "Reset Form"
     Else 'FLAG1 = FALSE/0
        Mouse.Animate Me, ""
        cmdAniLocal.Caption = "Set Form"
    End If
    Flag2 = False
    cmdAniGlobal.Caption = "Set System"

End Sub

Private Sub cmdBlock_Click()

    Mouse.Block = True
    tmrBlock.Enabled = True

End Sub

Private Sub cmdClip_Click()

  Static Clip As Boolean

    Clip = Not Clip
    If Clip Then
        Mouse.Clip.Control picClip.hWnd
        cmdClip.Caption = "Declip Cursor"
     Else 'CLIP = FALSE/0
        Mouse.Clip.Release
        cmdClip.Caption = "Clip Cursor"
    End If

End Sub

Private Sub cmdDeactivate_Click()

  Dim X As Long

    X = MsgBox("The mouse will be deactivated till this" & vbCrLf & "window-session, so You have to reboot." & vbCrLf & _
     "You want to continue?", vbExclamation Or vbYesNo)
    If X = vbYes Then
        Mouse.Deactivate
    End If

End Sub

Private Sub cmdEnd_Click()

    Unload Me

End Sub

Private Sub cmdHide_Click()

    Mouse.Visible = False
    tmrHide.Enabled = True
    cmdHide.Enabled = False
    Call DisplayInfos

End Sub

Private Sub cmdSetClick_Click()


    With Mouse
        .Position.X = Me.Left / Screen.TwipsPerPixelX + 10
        .Position.Y = Me.Top / Screen.TwipsPerPixelY + 10
        .Click (vbLeftButton)
    End With 'Mouse

End Sub

Private Sub cmdSpeedReset_Click()

    scrSpeed.Value = memSpeed
    scrDblClick.Value = memDblClick

End Sub

Private Sub cmdSwap_Click()

    Mouse.ButtonsSwapped = Not Mouse.ButtonsSwapped
    Call DisplayInfos

End Sub

Private Sub DisplayInfos()

    chkSwap.Value = -Mouse.ButtonsSwapped
    chkExists.Value = -Mouse.Exists
    chkWheel.Value = -Mouse.WheelExists
    chkVisible.Value = -Mouse.Visible
    lblButtons.Caption = Mouse.Buttons & " Buttons"

End Sub

Private Sub Form_Load()
pbar.Value = 50
    OldProc1 = GetWindowLong(Me.hWnd, GWL_WNDPROC)
    SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf TWndProc1 ' Subclass the entire form
    Set Mouse = New clsMouse
    With Mouse
        .WatchPosition = True
        .WatchSystemClicks = True
        memSpeed = .Speed
        scrSpeed.Value = .Speed
        memDblClick = .DoubleClickTime
        scrDblClick.Value = .DoubleClickTime
    End With 'Mouse
    Call DisplayInfos

End Sub

Private Sub Form_Unload(Cancel As Integer)


    With Mouse
        .Speed = memSpeed
        .DoubleClickTime = memDblClick
        .ButtonsSwapped = False
    End With 'Mouse
    SetWindowLong Me.hWnd, GWL_WNDPROC, OldProc1
    Set Mouse = Nothing

End Sub

Private Sub Mouse_PositionChanged()

    Me.Caption = "Mouse moved: " & Mouse.Position.X & " - " & Mouse.Position.Y

End Sub

Private Sub Mouse_SytemClick(ByVal Button As MouseButtonConstants)

    Me.Caption = "System Click: " & Button

End Sub

Private Sub scrDblClick_Change()

    Mouse.DoubleClickTime = scrDblClick.Value
    lblDblClick.Caption = Mouse.DoubleClickTime

End Sub

Private Sub scrSpeed_Change()

    Mouse.Speed = scrSpeed.Value
    lblSpeed.Caption = Mouse.Speed

End Sub

Private Sub tmrBlock_Timer()

    Mouse.Block = False
    tmrBlock.Enabled = False

End Sub

Private Sub tmrHide_Timer()

    Mouse.Visible = True
    tmrHide.Enabled = False
    cmdHide.Enabled = True
    Call DisplayInfos

End Sub

':)Roja's VB Code Fixer V1.1.92 (07.03.2004 07:57:24) 33 + 192 = 225 Lines Thanks Ulli for inspiration and lots of code.

