Attribute VB_Name = "HotKeys"
Public Const VK_ADD = &H6B
Public Const VK_ATTN = &HF6
Public Const VK_BACK = &H8
Public Const VK_CANCEL = &H3
Public Const VK_CAPITAL = &H14
Public Const VK_CLEAR = &HC
Public Const VK_CONTROL = &H11
Public Const VK_CRSEL = &HF7
Public Const VK_DECIMAL = &H6E
Public Const VK_DELETE = &H2E
Public Const VK_DIVIDE = &H6F
Public Const VK_DOWN = &H28
Public Const VK_END = &H23
Public Const VK_EREOF = &HF9
Public Const VK_ESCAPE = &H1B
Public Const VK_EXECUTE = &H2B
Public Const VK_EXSEL = &HF8
Public Const VK_F1 = &H70
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F2 = &H71
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_HELP = &H2F
Public Const VK_HOME = &H24
Public Const VK_INSERT = &H2D
Public Const VK_LBUTTON = &H1
Public Const VK_LCONTROL = &HA2
Public Const VK_LEFT = &H25
Public Const VK_LMENU = &HA4
Public Const VK_LSHIFT = &HA0
Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
Public Const VK_MENU = &H12
Public Const VK_MULTIPLY = &H6A
Public Const VK_NEXT = &H22
Public Const VK_NONAME = &HFC
Public Const VK_NUMLOCK = &H90
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_OEM_CLEAR = &HFE
Public Const VK_PA1 = &HFD
Public Const VK_PAUSE = &H13
Public Const VK_PLAY = &HFA
Public Const VK_PRINT = &H2A
Public Const VK_PRIOR = &H21
Public Const VK_PROCESSKEY = &HE5
Public Const VK_RBUTTON = &H2
Public Const VK_RCONTROL = &HA3
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_RMENU = &HA5
Public Const VK_RSHIFT = &HA1
Public Const VK_SCROLL = &H91
Public Const VK_SELECT = &H29
Public Const VK_SEPARATOR = &H6C
Public Const VK_SHIFT = &H10
Public Const VK_SNAPSHOT = &H2C
Public Const VK_SPACE = &H20
Public Const VK_SUBTRACT = &H6D
Public Const VK_TAB = &H9
Public Const VK_UP = &H26
Public Const VK_ZOOM = &HFB

Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function WaitMessage Lib "user32" () As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4
Public Const PM_REMOVE = &H1
Public Const PM_NOREMOVE = &H0
Public Const WM_HOTKEY = &H312

Public Const GWL_WNDPROC = (-4)

Public PrevProc As Long

Public Function KeyCallbacks(ByVal hWnd As Long, ByVal uMSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMSG = WM_HOTKEY Then
        Call ActOnKeys(wParam)
        wParam = 1
        Exit Function
    End If
    KeyCallbacks = CallWindowProc(PrevProc, hWnd, uMSG, wParam, lParam)
End Function

Public Sub ActOnKeys(ByVal KeyID As Byte)
    If KeyID = 0 Then
        Call frmMain.mnuLoadListFromFile_Click
    ElseIf KeyID = 1 Then
        Call frmMain.mnuPasteClipboard_Click
    ElseIf KeyID = 2 Then
        If Not frmMain.List1.ListCount = 0 Then
            Call frmMain.mnuSaveCurrList_Click
        End If
    ElseIf KeyID = 3 Then
        Call frmMain.mnuTrayHideMainWindow_Click
    ElseIf KeyID = 4 Then
        Call frmMain.mnuTrayShowMainWindow_Click
    End If
End Sub
