Attribute VB_Name = "ModMBKB"
Private Const K_B = 1024#
Private Const M_B = (K_B * 1024#) ' MegaBytes
Private Const G_B = (M_B * 1024#) ' GigaBytes
Private Const T_B = (G_B * 1024#) ' TeraBytes
Private Const P_B = (T_B * 1024#) ' PetaBytes
Private Const E_B = (P_B * 1024#) ' ExaBytes
Private Const Z_B = (E_B * 1024#) ' ZettaBytes
Private Const Y_B = (Z_B * 1024#) ' YottaBytes

Public Enum DISP_BYTES_FORMAT
    DISP_BYTES_LONG
    DISP_BYTES_SHORT
    DISP_BYTES_ALL
End Enum

Public Function FormatFileSize(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023 ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575 ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823# ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select
End Function

Public Function GetSizeBytes(Dec As Variant, Optional DispBytesFormat As DISP_BYTES_FORMAT = DISP_BYTES_ALL) As String
    Dim DispLong As String: Dim DispShort As String: Dim s As String
    If DispBytesFormat <> DISP_BYTES_SHORT Then DispLong = FormatNumber(Dec, 0) & " bytes" Else DispLong = ""
    If DispBytesFormat <> DISP_BYTES_LONG Then
        If Dec > Y_B Then
            DispShort = FormatNumber(Dec / Y_B, 2) & " YB"
        ElseIf Dec > Z_B Then
            DispShort = FormatNumber(Dec / Z_B, 2) & " ZB"
        ElseIf Dec > E_B Then
            DispShort = FormatNumber(Dec / E_B, 2) & " EB"
        ElseIf Dec > P_B Then
            DispShort = FormatNumber(Dec / P_B, 2) & " PB"
        ElseIf Dec > T_B Then
            DispShort = FormatNumber(Dec / T_B, 2) & " TB"
        ElseIf Dec > G_B Then
            DispShort = FormatNumber(Dec / G_B, 2) & " GB"
        ElseIf Dec > M_B Then
            DispShort = FormatNumber(Dec / M_B, 2) & " MB"
        ElseIf Dec > K_B Then
            DispShort = FormatNumber(Dec / K_B, 2) & " KB"
        Else
            DispShort = FormatNumber(Dec, 0) & " bytes"
        End If
    Else
        DispShort = ""
    End If
    Select Case DispBytesFormat
        Case DISP_BYTES_SHORT:
            GetSizeBytes = DispShort
        Case DISP_BYTES_LONG:
            GetSizeBytes = DispLong
        Case Else:
            GetSizeBytes = DispLong & " (" & DispShort & ")"
    End Select
End Function
