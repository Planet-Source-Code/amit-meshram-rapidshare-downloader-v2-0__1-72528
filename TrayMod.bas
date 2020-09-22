Attribute VB_Name = "TrayMod"
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 27 'Replace the szTip string's length with your tip's length
    'szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Private Const WM_MOUSEMOVE = &H200
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_RBUTTONDOWN = &H204

Sub CreateSysTrayIcon()
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hWnd = frmMain.Picture1.hWnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frmMain.ImageList1.ListImages.Item(1).ExtractIcon
    Tic.szTip = "Rapidshare Downloader v2.0"
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub

Sub CreateSysTrayIconAnim(ImgIndex As Integer)
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hWnd = frmMain.Picture1.hWnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = frmMain.ImageList1.ListImages.Item(ImgIndex).ExtractIcon
    Tic.szTip = "Rapidshare Downloader v2.0" & vbNullString
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
End Sub

Sub DeleteSysTrayIcon()
    Dim Tic As NOTIFYICONDATA, erg As Long
    Tic.cbSize = Len(Tic)
    Tic.hWnd = frmMain.Picture1.hWnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub
