Attribute VB_Name = "SystemMenu"
Option Explicit
Public Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function GetProp Lib "User32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetProp Lib "User32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GWL_WNDPROC  As Long = (-4)
Public Const MF_STRING = &H0
Public Const WM_SYSCOMMAND = &H112
Public Const MF_SEPARATOR = &H800

Public Const ID_ABOUT = 1000
Public Const ID_TRAY = 1001

Public Function HookFunc(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
  'this MUST be dimmed as the object passed!!!
   Dim obj As frmMain
   Dim foo As Long
   
   foo = GetProp(hWnd, "ObjectPointer")
   
  'Ignore "impossible" bogus case
   If (foo <> 0) Then
   
      CopyMemory obj, foo, 4
      On Error Resume Next
      HookFunc = obj.WindowProc(hWnd, msg, wp, lp)
      
      If (Err) Then
         UnhookWindow hWnd
         Debug.Print "Unhook on Error, #"; CStr(Err.Number)
         Debug.Print "  Desc: "; Err.Description
         Debug.Print "  Message, hWnd: &h"; Hex(hWnd), _
                             "Msg: &h"; Hex(msg), "Params:"; wp; lp
      End If

     'Make sure we don't get any foo->Release() calls
      foo = 0
      CopyMemory obj, foo, 4
   End If

End Function

Public Sub HookWindow(hWnd As Long, thing As Object)

   Dim foo As Long

   CopyMemory foo, thing, 4

   Call SetProp(hWnd, "ObjectPointer", foo)
   Call SetProp(hWnd, "OldWindowProc", GetWindowLong(hWnd, GWL_WNDPROC))
   Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf HookFunc)
   
End Sub

Public Sub UnhookWindow(hWnd As Long)
   
   Dim foo As Long

   foo = GetProp(hWnd, "OldWindowProc")
   
   If (foo <> 0) Then
      Call SetWindowLong(hWnd, GWL_WNDPROC, foo)
   End If
   
End Sub

Public Function InvokeWindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   InvokeWindowProc = CallWindowProc(GetProp(hWnd, "OldWindowProc"), hWnd, msg, wp, lp)
End Function

