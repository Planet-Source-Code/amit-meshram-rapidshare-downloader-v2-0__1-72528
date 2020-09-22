VERSION 5.00
Begin VB.UserControl Registry 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   DrawWidth       =   56
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
End
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegSaveKey Lib "advapi32.dll" Alias "RegSaveKeyA" (ByVal hKey As Long, ByVal lpFile As String, lpSecurityAttributes As Any) As Long 'notification Constants
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type SECURITY_ATTRIBUTES
     nLength As Long
     lpSecurityDescriptor As Long
     bInheritHandle As Long
End Type

Private Enum NotifyFilter
    REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
    REG_NOTIFY_CHANGE_LAST_SET = &H4
    REG_NOTIFY_CHANGE_NAME = &H1
    REG_NOTIFY_CHANGE_SECURITY = &H8
End Enum

Public Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Public Enum hKeySave
    sHKEY_CLASSES_ROOT = &H80000000
    sHKEY_CURRENT_USER = &H80000001
    sHKEY_LOCAL_MACHINE = &H80000002
    sHKEY_USERS = &H80000003
    sHKEY_CURRENT_CONFIG = &H80000005
End Enum

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Const ERROR_NO_MORE_ITEMS = 259

Public Enum REG_TYPE
 REG_SZ = 1                         ' Unicode nul terminated string D
 REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
 REG_BINARY = 3                     ' Free form binary
 REG_DWORD = 4                      ' 32-bit number D
 REG_DWORD_LITTLE_ENDIAN = 4        ' 32-bit number (same as REG_DWORD) D
 REG_MULTI_SZ = 7                   ' Multiple Unicode strings
End Enum

Private Const ERROR_SUCCESS = 0

Private Const REG_OPTION_VOLATILE = 1
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore

Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const REG_CREATED_NEW_KEY = 1 ' had to create a key
Private Const REG_OPENED_EXISTING_KEY = 2 'opened a key only

Public Event EnumKey(strComputerName As String, strHive As String, strKey As String, strSubKeys As String)
Public Event EnumValue(strComputerName As String, lngHKey As String, strSubKey As String, strValueName As String, strActualValue As String)
Public Event Error(lngErrorNumber As Long, strDescription As String)

Public Function SaveValueEx(strComputerName As String, lngHKey As hKeySave, strSubKey As String, strValueName As String, strData As String, lngDataType As REG_TYPE) As Boolean
    Dim lngResult As Long ' holds the handle of new subkey
    Dim lngRet As Long ' holds returned values
    Dim lngCreatedOROpened As Long
    
    If strComputerName <> "" Then
        lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
        If lngRet <> 0 Then
            RaiseEvent Error(lngRet, Err.Description)
            Exit Function
        End If
    End If
    
    If strComputerName <> "" Then
        lngRet = RegCreateKeyEx(lngResult, strSubKey, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal &O0, lngResult, lngCreatedOROpened)
        If lngRet <> 0 Then
            RaiseEvent Error(lngRet, Err.Description)
            Exit Function
        End If
    Else
        lngRet = RegCreateKeyEx(lngHKey, strSubKey, 0, 0, REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal &O0, lngResult, lngCreatedOROpened)
    End If
    
    If lngRet <> ERROR_SUCCESS Then
        RaiseEvent Error(lngRet, Err.Description)
        Exit Function
    End If
    
    If lngCreatedOROpened = REG_CREATED_NEW_KEY Then
    ElseIf lngCreatedOROpened = REG_OPENED_EXISTING_KEY Then
    End If
    
    Select Case lngDataType 'VarType(strdata)
    
    Case REG_SZ 'WORKS 'vbString '
        
        lngRet = RegSetValueEx(lngResult, strValueName, 0, REG_SZ, ByVal strData, Len(strData))
    
    
    Case REG_MULTI_SZ 'WORKS
        
        lngRet = RegSetValueEx(lngResult, strValueName, 0, REG_MULTI_SZ, ByVal strData, Len(strData))
    
    Case REG_BINARY 'WORKS
        
        Dim lngLenstrData As Long
        Dim bytArray() As Byte
        
        lngLenstrData = Len(strData)
        ReDim bytArray(lngLenstrData)
        
        Dim I As Integer
        For I = 1 To lngLenstrData
            bytArray(I) = Asc(Mid$(strData, I, 1))
        Next
        
        lngRet = RegSetValueEx(lngResult, strValueName, 0, REG_BINARY, bytArray(1), lngLenstrData)
    
    Case REG_DWORD 'WORKS
        lngRet = RegSetValueEx(lngResult, strValueName, 0, REG_DWORD, CLng(strData), 4)
    Case REG_DWORD_LITTLE_ENDIAN 'WORKS
        lngRet = RegSetValueEx(lngResult, strValueName, 0, REG_DWORD_LITTLE_ENDIAN, CLng(strData), 4)
    Case REG_EXPAND_SZ
        lngRet = RegSetValueEx(lngResult, strValueName, 0, REG_EXPAND_SZ, ByVal strData, Len(strData))
    Case Else
    End Select
    If lngRet <> ERROR_SUCCESS Then
        RaiseEvent Error(lngRet, Err.Description)
        Exit Function
    End If
    
    lngRet = RegCloseKey(lngResult)
    
    If lngRet <> ERROR_SUCCESS Then
        RaiseEvent Error(lngRet, Err.Description)
        Exit Function
    Else
        SaveValueEx = True
    End If
End Function

Public Function DeleteKeyEx(strComputerName As String, lngHKey As hKey, strSubKey As String) As Boolean
    Dim lngResult As Long ' holds the handle of new subkey
    Dim lngRet As Long ' holds returned values
    If strComputerName <> "" Then
        lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
    End If
    
    If strComputerName <> "" Then
        lngRet = RegOpenKeyEx(lngResult, strSubKey, 0, KEY_ALL_ACCESS, lngResult)
    Else
        lngRet = RegOpenKeyEx(lngHKey, strSubKey, 0, KEY_ALL_ACCESS, lngResult)
    End If
    
    If lngRet = 0 Then 'it does
        lngRet = RegDeleteKey(lngResult, "")
        If lngRet = 0 Then
            DeleteKeyEx = True
        End If
    End If
End Function


Public Function DeleteValueEx(strComputerName As String, lngHKey As hKey, strSubKey As String, strValue As String) As Boolean
    Dim lngResult As Long ' holds the handle of new subkey
    Dim lngRet As Long ' holds returned values
    If strComputerName <> "" Then
        lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
    End If
    
    If strComputerName <> "" Then
        lngRet = RegOpenKeyEx(lngResult, strSubKey, 0, KEY_ALL_ACCESS, lngResult)
    Else
        lngRet = RegOpenKeyEx(lngHKey, strSubKey, 0, KEY_ALL_ACCESS, lngResult)
    End If
    
    If lngRet <> ERROR_SUCCESS Then
        RaiseEvent Error(lngRet, Err.Description)
        Exit Function
    End If

    lngRet = RegDeleteValue(lngResult, strValue)
    
    If lngRet <> ERROR_SUCCESS Then
        Exit Function
    Else
        DeleteValueEx = True
    End If
    RegCloseKey lngResult
End Function

Public Sub EnumerateKey(strComputerName As String, lngHKey As hKey, strSubKey As String)
    Dim lngResult As Long ' holds the handle of new subkey
    Dim lngRet As Long ' holds returned values
    Dim strBuffer As String
    Dim lngBufferLength As Long
    Dim lngIndex As Long ' holds the index of the current Value
    
    Dim ft As FILETIME

    If strComputerName <> "" Then
        lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
        
        If lngRet > 0 Then 'no connection
            RaiseEvent Error(lngRet, Err.Description)
            Exit Sub
        End If
        
    End If
    If strComputerName <> "" Then
        lngRet = RegOpenKeyEx(lngResult, strSubKey, 0, KEY_READ, lngResult)
    Else
        lngRet = RegOpenKeyEx(lngHKey, strSubKey, 0, KEY_READ, lngResult)
    End If
        
    Do While lngRet = ERROR_SUCCESS
        lngBufferLength = 2000
        strBuffer = String$(lngBufferLength, 0)
    
        If lngRet <> ERROR_SUCCESS Then
                RaiseEvent Error(lngRet, Err.Description)
            Exit Sub
        End If
        
        lngRet = RegEnumKeyEx(lngResult, lngIndex, strBuffer, lngBufferLength, 0, vbNullString, lngRet, ft)
        lngIndex = lngIndex + 1
        
        If lngRet = ERROR_NO_MORE_ITEMS Then
            GoTo EndLoop
        End If
        
        Select Case lngHKey
        
        Case HKEY_LOCAL_MACHINE
            
            RaiseEvent EnumKey(strComputerName, "HKEY_LOCAL_MACHINE", strSubKey, Left$(strBuffer, lngBufferLength))
        
        Case HKEY_CURRENT_USER
            
            RaiseEvent EnumKey(strComputerName, "HKEY_CURRENT_USER", strSubKey, Left$(strBuffer, lngBufferLength))
        
        Case HKEY_CLASSES_ROOT
            
            RaiseEvent EnumKey(strComputerName, "HKEY_CLASSES_ROOT", strSubKey, Left$(strBuffer, lngBufferLength))
        
        Case HKEY_USERS
            
            RaiseEvent EnumKey(strComputerName, "HKEY_USERS", strSubKey, Left$(strBuffer, lngBufferLength))
        
        Case HKEY_CURRENT_CONFIG
            
            RaiseEvent EnumKey(strComputerName, "HKEY_CURRENT_CONFIG", strSubKey, Left$(strBuffer, lngBufferLength))
        
        Case HKEY_DYN_DATA
            RaiseEvent EnumKey(strComputerName, "HKEY_DYN_DATA", strSubKey, Left$(strBuffer, lngBufferLength))
        
        End Select
    Loop
EndLoop:
    RegCloseKey lngResult
End Sub

Public Function GetValueEx(strComputerName As String, lngHKey As hKey, strSubKey As String, strValueName As Variant) As Variant
    Dim lngResult As Long ' holds the handle of new subkey
    Dim lngRet As Long ' holds returned values
    Dim strBuffer As String
    Dim lngBufferLength As Long
    Dim lngReturn As Long
    Dim intData As Integer

    
    lngBufferLength = 4000 ' this is the largest value the registry can hold
    strBuffer = String$(lngBufferLength, 0)
    If strComputerName <> "" Then
        lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
    End If
    
    If lngRet = ERROR_SUCCESS Then 'all went ok
        
        If strComputerName <> "" Then
            lngRet = RegOpenKeyEx(lngResult, strSubKey, 0, KEY_READ, lngResult)
        Else
            lngRet = RegOpenKeyEx(lngHKey, strSubKey, 0, KEY_READ, lngResult)
        End If
            
            If lngRet <> ERROR_SUCCESS Then
                GetValueEx = "NO DATA"
                Exit Function
            End If
            
            Dim lngValueType As Long
            lngRet = RegQueryValueEx(lngResult, strValueName, 0, lngValueType, ByVal 0, lngBufferLength) 'ByVal 0)
            
            If lngRet <> ERROR_SUCCESS Then
                GetValueEx = "NO DATA"
                Exit Function
            End If
            
            Select Case lngValueType
                Case REG_SZ, REG_MULTI_SZ ' a string
                    strBuffer = String$(lngBufferLength, Chr$(0))
                    
                    lngRet = RegQueryValueEx(lngResult, strValueName, 0, lngValueType, ByVal strBuffer, lngBufferLength)
                    GetValueEx = Left$(strBuffer, lngBufferLength - 1)
                Case REG_BINARY ' Free form binary
                    Dim MyBuffer() As Byte
                    ReDim MyBuffer(lngBufferLength - 1) As Byte
                    lngRet = RegQueryValueEx(lngResult, strValueName, 0&, lngValueType, MyBuffer(0), lngBufferLength)
                    GetValueEx = MyBuffer
                Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN '32-bit unsigned integer
                    lngRet = RegQueryValueEx(lngResult, strValueName, 0, 0, intData, lngBufferLength)
                    If lngRet = ERROR_SUCCESS Then
                        GetValueEx = intData
                    End If
                Case REG_EXPAND_SZ
                    Dim resBinary() As Byte
                    Dim Length As Long
                    Dim resString As String
                    Length = 1024
                    ReDim resBinary(0 To Length - 1)
                    lngRet = RegQueryValueEx(lngResult, strValueName, 0, lngValueType, resBinary(0), Length)  'ByVal 0)
                    If Length <> 0 Then
                        resString = Space$(Length - 1)
                        CopyMemory ByVal resString, resBinary(0), Length - 1
                        Dim s$, dl&
                        Dim y As String * 500
                        s$ = resString
                        dl& = ExpandEnvironmentStrings(s$, y, 499)
                        GetValueEx = y
                    End If
            End Select
            RegCloseKey lngResult
    Else
       GetValueEx = "NO DATA"
    End If
    
End Function


Public Sub EnumerateValues(strComputerName As String, lngHKey As hKey, ByVal strSubKey As String)
    
    Dim intIndex As Integer
    Dim strBuffer As String
    Dim lngBufferLength As String
    Dim lngRet As Long
    Dim lngResult As Long
    
    Dim bytData As Byte
    Dim lngType As Long
    Dim lngData As Long



If strComputerName <> "" Then
    lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
    
    If lngRet <> ERROR_SUCCESS Then
        'an error
        RaiseEvent Error(lngRet, Err.Description)
        Exit Sub
    End If
    
End If
If strComputerName <> "" Then
    lngRet = RegOpenKeyEx(lngResult, strSubKey, 0, KEY_ALL_ACCESS, lngResult)
    If lngRet <> ERROR_SUCCESS Then
        RaiseEvent Error(lngRet, Err.Description)
        Exit Sub
    End If
Else
    lngRet = RegOpenKeyEx(lngHKey, strSubKey, 0, KEY_ALL_ACCESS, lngResult)
    
    If lngRet <> ERROR_SUCCESS Then
        RaiseEvent Error(lngRet, Err.Description)
        Exit Sub
    End If
End If
Do While lngRet = ERROR_SUCCESS

    lngBufferLength = 255
    strBuffer = String$(lngBufferLength, 0)
    
    Dim lngLength As Long
    Dim strActualValue As String
    
    lngRet = RegEnumValue(lngResult, intIndex, strBuffer, 255, 0, ByVal 0&, ByVal 0&, ByVal 0&)
    
    If lngRet = ERROR_NO_MORE_ITEMS Then GoTo EndLoop
    
    If lngRet = ERROR_SUCCESS Then
        Select Case lngHKey
        
            Case HKEY_LOCAL_MACHINE
                    strActualValue = GetValueEx(strComputerName, HKEY_LOCAL_MACHINE, strSubKey, Trim$(strBuffer))
                    RaiseEvent EnumValue(strComputerName, "HKEY_LOCAL_MACHINE", strSubKey, Trim$(strBuffer), strActualValue)
        
            Case HKEY_CURRENT_USER
                    strActualValue = GetValueEx(strComputerName, HKEY_CURRENT_USER, strSubKey, Trim$(strBuffer))
                    RaiseEvent EnumValue(strComputerName, "HKEY_CURRENT_USER", strSubKey, Trim$(strBuffer), strActualValue)
            
            Case HKEY_CLASSES_ROOT
                    strActualValue = GetValueEx(strComputerName, HKEY_CLASSES_ROOT, strSubKey, Trim$(strBuffer))
                    RaiseEvent EnumValue(strComputerName, "HKEY_CLASSES_ROOT", strSubKey, Trim$(strBuffer), strActualValue)
            
            Case HKEY_CURRENT_CONFIG
                    strActualValue = GetValueEx(strComputerName, HKEY_CURRENT_CONFIG, strSubKey, Trim$(strBuffer))
                    RaiseEvent EnumValue(strComputerName, "HKEY_CURRENT_CONFIG", strSubKey, Trim$(strBuffer), strActualValue)
            
            Case HKEY_DYN_DATA
                    strActualValue = GetValueEx(strComputerName, HKEY_DYN_DATA, strSubKey, Trim$(strBuffer))
                    RaiseEvent EnumValue(strComputerName, "HKEY_DYN_DATA", strSubKey, Trim$(strBuffer), strActualValue)
                    
            Case HKEY_USERS
                    strActualValue = GetValueEx(strComputerName, HKEY_USERS, strSubKey, Trim$(strBuffer))
                    RaiseEvent EnumValue(strComputerName, "HKEY_USERS", strSubKey, Trim$(strBuffer), strActualValue)
                    
        End Select
    
    Else
        RaiseEvent Error(lngRet, Err.Description)
    End If
    intIndex = intIndex + 1
EndLoop:
Loop
    RegCloseKey lngResult
End Sub

Public Sub CreateKeyEx(strComputerName As String, lngHKey As hKey, strSubKey As String)
    Dim lngRet As Long
    Dim lngCreatedORExisted As Long
    Dim lngResult As Long
    
    If strComputerName <> "" Then
        lngRet = RegConnectRegistry("\\" & strComputerName, lngHKey, lngResult)
        
        If RegCreateKeyEx(lngResult, strSubKey, ByVal 0&, "", REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal 0&, lngResult, lngCreatedORExisted) <> ERROR_SUCCESS Then
            RaiseEvent Error(lngRet, Err.Description)
        End If
    Else ' no computer name given
        
        If RegCreateKeyEx(lngHKey, strSubKey, ByVal 0&, "", REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal 0&, lngResult, lngCreatedORExisted) <> ERROR_SUCCESS Then
            RaiseEvent Error(lngRet, Err.Description)
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
End Sub
