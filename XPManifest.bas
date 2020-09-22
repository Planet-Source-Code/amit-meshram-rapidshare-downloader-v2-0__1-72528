Attribute VB_Name = "XPManifest"
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Public Sub InitXP()
    Dim strPath As String
    Dim strData As String
    Dim FF As Integer
    
    On Error Resume Next
    
    strPath = App.Path & IIf(Right(App.Path, 1) = "\", vbNullString, "\")
    strPath = strPath & App.EXEName & IIf(LCase(Right(App.EXEName, 4)) = ".exe", ".manifest", ".exe.manifest")
    
    If Dir(strPath, vbReadOnly Or vbSystem Or vbHidden) <> vbNullString Then GoTo InitControls
    strData = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>" & vbCrLf
    strData = strData & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">" & vbCrLf
    strData = strData & "     <assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & "HybridDesign.WindowsXP.Example" & Chr(34) & " type=" & Chr(34) & "win32" & Chr(34) & " />" & vbCrLf
    strData = strData & "     <description>Windows XP Theme.</description>" & vbCrLf
    strData = strData & "     <dependency>" & vbCrLf
    strData = strData & "          <dependentAssembly>" & vbCrLf
    strData = strData & "               <assemblyIdentity type=" & Chr(34) & "win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " />" & vbCrLf
    strData = strData & "          </dependentAssembly>" & vbCrLf
    strData = strData & "     </dependency>" & vbCrLf
    strData = strData & "</assembly>"
    
    FF = FreeFile
    Open strPath For Output As #FF
        Print #FF, strData
    Close #FF
    
    SetAttr strPath, vbHidden Or vbSystem Or vbReadOnly Or vbArchive
InitControls:
    Call InitCommonControls
End Sub

