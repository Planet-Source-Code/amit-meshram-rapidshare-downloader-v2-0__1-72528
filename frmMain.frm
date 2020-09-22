VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Rapidshare Downloader v2.0"
   ClientHeight    =   7515
   ClientLeft      =   1425
   ClientTop       =   1215
   ClientWidth     =   8220
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8220
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   4320
   End
   Begin VB.TextBox txtClip2 
      Height          =   675
      Left            =   1020
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   630
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtClipboard 
      Height          =   675
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   630
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   540
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   31
      Top             =   1950
      Visible         =   0   'False
      Width           =   1035
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   180
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":ADA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":11606
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":17E68
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E6CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24F2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B78E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":31FF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":38852
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3F0B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":45916
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4BBB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":52412
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin RSD.Registry Reg1 
      Left            =   -300
      Top             =   -240
      _extentx        =   847
      _extenty        =   847
   End
   Begin VB.Timer tmrTray 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1380
      Top             =   6000
   End
   Begin VB.Timer tmrFile 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1440
      Top             =   4770
   End
   Begin VB.TextBox txtCurrFile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   6780
      Width           =   6630
   End
   Begin VB.TextBox txtFilePath 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1590
      TabIndex        =   27
      Top             =   3360
      Width           =   5430
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse"
      Height          =   345
      Left            =   7095
      TabIndex        =   26
      Top             =   3360
      Width           =   1005
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   7530
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "frmMain.frx":58C74
      Left            =   120
      List            =   "frmMain.frx":58C76
      TabIndex        =   23
      Top             =   510
      Width           =   7995
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   3840
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   8400
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9000
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel All Download"
      Height          =   390
      Left            =   5940
      TabIndex        =   12
      Top             =   4905
      Width           =   2145
   End
   Begin VB.CommandButton CmdDownload 
      Caption         =   "Start All Download"
      Height          =   390
      Left            =   3480
      TabIndex        =   11
      Top             =   4905
      Width           =   2400
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   3495
      TabIndex        =   9
      Top             =   3810
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StPanel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   7170
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8740
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "PM 06:20"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Text            =   "http://rapidshare.com/files/180564572/7zip_www.sxforum.org.rar"
      Top             =   120
      Width           =   5430
   End
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   9570
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Caption         =   "Current File"
      Height          =   255
      Left            =   150
      TabIndex        =   30
      Top             =   6810
      Width           =   1185
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   150
      TabIndex        =   28
      Top             =   6420
      Width           =   7935
   End
   Begin VB.Label Label7 
      Caption         =   "Download Folder"
      Height          =   255
      Left            =   150
      TabIndex        =   25
      Top             =   3390
      Width           =   1305
   End
   Begin VB.Label LblWait2 
      Height          =   255
      Left            =   5310
      TabIndex        =   22
      Top             =   5355
      Width           =   2775
   End
   Begin VB.Label lblWait1 
      Height          =   255
      Left            =   2790
      TabIndex        =   21
      Top             =   5355
      Width           =   2460
   End
   Begin VB.Label Label8 
      Caption         =   "Total Time"
      Height          =   240
      Left            =   135
      TabIndex        =   20
      Top             =   6060
      Width           =   1125
   End
   Begin VB.Label lblTakeTime 
      Caption         =   "0 Days, 0 Hours, 0 Minutes and 0 Seconds."
      Height          =   255
      Left            =   2025
      TabIndex        =   19
      Top             =   6060
      Width           =   6060
   End
   Begin VB.Label lblTime 
      Caption         =   "0 Days, 0 Hours, 0 Minutes and 0 Seconds."
      Height          =   255
      Left            =   2025
      TabIndex        =   18
      Top             =   5700
      Width           =   6060
   End
   Begin VB.Label Label5 
      Caption         =   "Time Remaining"
      Height          =   240
      Left            =   135
      TabIndex        =   17
      Top             =   5700
      Width           =   1755
   End
   Begin VB.Label lblRapidStatus 
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Top             =   4530
      Width           =   4590
   End
   Begin VB.Label Label3 
      Caption         =   "Speed"
      Height          =   240
      Left            =   150
      TabIndex        =   15
      Top             =   4980
      Width           =   1125
   End
   Begin VB.Label lblSpeed 
      Caption         =   "In KBPS"
      Height          =   240
      Left            =   2010
      TabIndex        =   14
      Top             =   4980
      Width           =   1365
   End
   Begin VB.Label lblWait 
      Caption         =   "Wait : "
      Height          =   240
      Left            =   150
      TabIndex        =   13
      Top             =   5370
      Width           =   2565
   End
   Begin VB.Label lblPercentage 
      Caption         =   "Percent % Completed..."
      Height          =   240
      Left            =   3510
      TabIndex        =   10
      Top             =   4170
      Width           =   4575
   End
   Begin VB.Label lblRemaining 
      Caption         =   "In KB/MB"
      Height          =   285
      Left            =   2010
      TabIndex        =   8
      Top             =   4590
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "File Remaining"
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   4590
      Width           =   1125
   End
   Begin VB.Label lblSaved 
      Caption         =   "In KB/MB"
      Height          =   285
      Left            =   2010
      TabIndex        =   6
      Top             =   4200
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "File Saved"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label lblSize 
      Caption         =   "In KB/MB"
      Height          =   285
      Left            =   2010
      TabIndex        =   3
      Top             =   3810
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "File Size"
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   3810
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Rapidshare URL"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1815
   End
   Begin VB.Menu mnu1 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuLoadListFromFile 
         Caption         =   "&Load List from File"
      End
      Begin VB.Menu mnuSaveCurrList 
         Caption         =   "&Save Current List"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClearAllList 
         Caption         =   "&Clear All List"
      End
      Begin VB.Menu sap1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPasteClipboard 
         Caption         =   "&Paste Links From Clipboard"
      End
      Begin VB.Menu mnuDeleteCurrFile 
         Caption         =   "&Delete Current File"
      End
      Begin VB.Menu sap2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "&Move Up"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "&Move Down"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnu2LoadListFromFile 
         Caption         =   "&Load List From File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuPasteLinksClipboard 
         Caption         =   "&Paste Links from Clipboard"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShowMainWindow 
         Caption         =   "&Show Main Window"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuTrayHideMainWindow 
         Caption         =   "&Hide Main Window"
         Shortcut        =   ^X
      End
      Begin VB.Menu TraySap1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayStartDownloding 
         Caption         =   "&Start Downloading"
      End
      Begin VB.Menu mnuTrayStopDownloading 
         Caption         =   "&Stop Downloading"
      End
      Begin VB.Menu TraySap2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SH As New Shell
Dim SHFF As Folder

Dim RetHotKey1, RetHotKey2, RetHotKey3, RetHotKey4, RetHotKey5 As Boolean
Dim icoIndex As Integer

Sub TerminateTimer()
    Timer1.Enabled = False
End Sub

Sub ResetAllControl()
    PB1.Value = 0
    lblPercentage.Caption = "Percent % Completed..."
    lblRapidStatus.Caption = ""
    lblSize = "In KB/MB"
    lblSaved = "In KB/MB"
    lblRemaining = "In KB/MB"
    lblSpeed = "in KBPS"
    lblWait = "Wait : "
    lblWait1 = ""
    LblWait2 = ""
    lblTime = "0 Days, 0 Hours, 0 Minutes and 0 Seconds."
    lblTakeTime = "0 Days, 0 Hours, 0 Minutes and 0 Seconds."
    txtCurrFile = ""
    Label9 = ""
End Sub

Private Sub CmdAdd_Click()
    If txtURL = "" Then Exit Sub
    If InStr(txtURL.Text, "rapidshare.com") Then
        List1.AddItem txtURL.Text
        txtURL.Text = ""
    Else
        MsgBox "Please Use Only Rapidshare Links!!!"
    End If
End Sub

Sub CmdCancel_Click()
    Inet1.Cancel
    Inet2.Cancel
    Inet3.Cancel
    
    If frmMain.Tag = "Cancel" Then
        Inet1.Cancel
        Inet2.Cancel
        Inet3.Cancel
    End If
    Timer1.Enabled = False
    lblWait.Caption = "Action : Cancelled"
    Cnt30s = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msg
msg = "Do you really want to exit the application?"
If MsgBox(msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
Call RestoreHook
End Sub

Private Sub mnuTrayStartDownloding_Click()
    If Not List1.ListIndex = -1 And bDone = True Then
        Call GetInfo1(Inet1, List1.List(0))
    End If
End Sub

Private Sub tmrFile_Timer()
    If Not List1.ListCount = -1 And bDone = True Then
        Call GetInfo1(Inet1, List1.List(0))
    End If
    tmrFile.Enabled = False
End Sub

Sub ProcessListForDownload(URLLink As String)
    Dim lst1 As Integer
    Dim ls2() As String
    TotalFileToDownload = List1.ListCount
    FileDownloderCounter = 0
    Call GetInfo1(Inet1, URLLink)
End Sub

Private Sub CmdDownload_Click()
Dim I As Integer
    Inet1.Cancel
    If frmMain.Tag = "Cancel" Then
        Inet1.Cancel
    End If
    If Not List1.ListCount = -1 Then
        Call GetInfo1(Inet1, List1.List(0))
    End If
End Sub

Private Sub CmdBrowse_Click()
    On Error Resume Next
    Set SHFF = SH.BrowseForFolder(hWnd, "Choose Folder you want to save your downloaded file.", 1)
    With SHFF.Items.Item
        txtFilePath.Text = .Path
    End With
    SaveBrowsedPath
    GetBrowsedPath
End Sub

Private Sub Form_Load()
    InitXP
    If App.PrevInstance = True Then
        Unload Me
        End
    End If
    
    bDone = False
    txtFilePath.Text = GetBrowsedPath
    Call CreateSysTrayIcon
    Call RegHotKeys(Me.hWnd)
    Call MakeSystemMenu
End Sub

Private Function SaveBrowsedPath() As Boolean
    If Not txtFilePath.Text = "" Then
        Reg1.CreateKeyEx "", HKEY_LOCAL_MACHINE, "Software\RSD"
        Reg1.SaveValueEx "", sHKEY_LOCAL_MACHINE, "Software\RSD", "Path", txtFilePath.Text, REG_SZ
        SaveBrowsedPath = True
    End If
End Function

Private Function GetBrowsedPath() As String
    Dim nRet  As String
    nRet = Reg1.GetValueEx("", HKEY_LOCAL_MACHINE, "Software\RSD", "Path")
    If nRet = "NO DATA" Then
        GetBrowsedPath = App.Path
    Else
        GetBrowsedPath = nRet
    End If
End Function

Private Function CheckForFileList() As Boolean
    If List1.ListCount = 0 Then
        CheckForFileList = False
    Else
        CheckForFileList = True
    End If
End Function

Private Function CountList() As Integer
    CountList = List1.ListCount
    TotalFileToDownload = List1.ListCount
End Function

Private Sub Form_Resize()
    If Not Me.WindowState = vbMinimized Then
        Me.Width = 8295 + 150
        Me.Height = 7950 + 150
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DeleteSysTrayIcon
    Call UnRegHotKeys(Me.hWnd)
    Call RestoreHook
End Sub

Private Sub List1_Click()
    If List1.SelCount > 1 Then
        mnuMoveDown.Enabled = False
        mnuMoveUp.Enabled = False
        Exit Sub
    Else
        mnuMoveDown.Enabled = True
        mnuMoveUp.Enabled = True
    End If
    If List1.Selected(0) Then
        mnuMoveUp.Enabled = False
    Else
        mnuMoveUp.Enabled = True
    End If
    If List1.Selected(List1.ListCount - 1) Then
        mnuMoveDown.Enabled = False
    Else
        mnuMoveDown.Enabled = True
    End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And CheckForFileList = True Then
        PopupMenu mnu1, vbPopupMenuLeftAlign
    End If
    
    If Button = 2 And CheckForFileList = False Then
        If CheckForClipboard = False Then
            mnuPasteLinksClipboard.Enabled = False
        Else
            mnuPasteLinksClipboard.Enabled = True
        End If
        PopupMenu mnuPopup2, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub mnu2LoadListFromFile_Click()
    Dim ofn As OPENFILENAME
    Dim nRet As Long
    Dim Filter As String
    
    ofn.lStructSize = Len(ofn)
    ofn.hInstance = App.hInstance
    ofn.hwndOwner = Me.hWnd
    
    Filter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    'ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "Open File.."
    ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    
    nRet = GetOpenFileName(ofn)
    
    If nRet Then
        Open ofn.lpstrFile For Input As #1
            Do While Not EOF(1)
                Line Input #1, strData
                If InStr(strData, "rapidshare.com") Then
                    List1.AddItem Trim(strData)
                End If
            Loop
        Close #1
    End If
End Sub

Private Sub mnuClearAllList_Click()
    List1.Clear
End Sub

Sub mnuDeleteCurrFile_Click()
Dim I As Integer
If List1.ListIndex = -1 Then Exit Sub
    I = List1.ListIndex
    List1.RemoveItem List1.ListIndex
    List1.ListIndex = I - 1
End Sub

Sub mnuLoadListFromFile_Click()
    Dim ofn As OPENFILENAME
    Dim nRet As Long
    Dim Filter As String
    
    ofn.lStructSize = Len(ofn)
    ofn.hInstance = App.hInstance
    ofn.hwndOwner = Me.hWnd
    
    Filter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrTitle = "Open File.."
    ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    
    nRet = GetOpenFileName(ofn)
    
    If nRet Then
        Open ofn.lpstrFile For Input As #1
            Do While Not EOF(1)
                Line Input #1, strData
                If InStr(strData, "rapidshare.com") Then
                    List1.AddItem Trim(strData)
                End If
            Loop
        Close #1
    End If
End Sub

Private Sub mnuMoveDown_Click()
Dim strTemp As String
Dim Count As Integer
Count = List1.ListIndex
If Count > -1 Then
    strTemp = List1.List(Count)
    List1.AddItem strTemp, (Count + 2)
    List1.RemoveItem (Count)
    List1.Selected(Count + 1) = True
End If
End Sub

Private Sub mnuMoveUp_Click()
Dim strTemp As String
Dim Count As Integer
Count = List1.ListIndex
If Count > -1 Then
    strTemp = List1.List(Count)
    List1.AddItem strTemp, (Count - 1)
    List1.RemoveItem (Count + 1)
    List1.Selected(Count - 1) = True
End If
End Sub

Function CheckForClipboard() As Boolean
    Dim str1 As String
    Dim strData As String
    
    str1 = Clipboard.GetText(vbCFText)
    txtClip2.Text = str1
    
    Open App.Path & "\paste2.tmp" For Output As #1
        Print #1, Trim(txtClip2.Text)
    Close #1
    Open App.Path & "\paste2.tmp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strData
            If InStr(strData, "http://rapidshare.com") Then
                CheckForClipboard = True
            Else
                CheckForClipboard = False
            End If
        Loop
    Close #1
    txtClip2.Text = ""
    Kill App.Path & "\paste2.tmp"
End Function

Sub mnuPasteClipboard_Click()
    Dim str1 As String
    Dim strData As String
    
    str1 = Clipboard.GetText(vbCFText)
    txtClipboard.Text = str1
    
    Open App.Path & "\paste.tmp" For Output As #1
        Print #1, Trim(txtClipboard.Text)
    Close #1
    
    Open App.Path & "\paste.tmp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strData
            If InStr(strData, "http://rapidshare.com") Then
                List1.AddItem strData
            End If
        Loop
    Close #1
    Kill App.Path & "\paste.tmp"
    txtClipboard.Text = ""
End Sub

Private Sub mnuPasteLinksClipboard_Click()
Call mnuPasteClipboard_Click
End Sub

Sub mnuSaveCurrList_Click()
    Dim ofn As OPENFILENAME
    Dim nRet As Long
    Dim Filter As String
    
    Filter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Me.hWnd
    ofn.hInstance = App.hInstance
    
    ofn.lpstrFilter = Filter
    ofn.lpstrDefExt = ".txt"
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = App.Path
    ofn.lpstrTitle = "Save File.."
    ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    
    nRet = GetSaveFileName(ofn)

    If nRet Then
        Open ofn.lpstrFile For Output As #1
            For I = 0 To List1.ListCount - 1
                Print #1, Trim(List1.List(I))
            Next
        Close #1
    End If
End Sub

Private Sub mnuTrayExit_Click()
    Unload Me
End Sub

Sub mnuTrayHideMainWindow_Click()
    Me.Hide
End Sub

Sub mnuTrayShowMainWindow_Click()
    Me.Show
End Sub

Private Sub mnuTrayStopDownloading_Click()
    Call CmdCancel_Click
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x
        Case Is = WM_LBUTTONDOWN
            If frmMain.Visible Then frmMain.Visible = False Else frmMain.Visible = True
        Case Is = WM_RBUTTONDOWN
            PopupMenu mnuTray
    End Select
End Sub

Private Sub Timer1_Timer()
    Cnt = Cnt - 1
    lblWait.Caption = "Download Will Start In : " & Cnt & " Seconds"
    If Cnt = -1 Then: lblWait.Caption = "File Downloading Started..."
    If Cnt = 0 Then
        lblWait.Caption = "File Downloading Started..."
        Call DownloadCreate
        Timer1.Enabled = False
        Cnt = 0
    End If
End Sub

Private Sub Inet3_StateChanged(ByVal State As Integer)
    StPanel.Panels(1).Text = GetStatus(State, Inet3)
End Sub

'==========================================================================================================
'Process HOT KEYS
'==========================================================================================================
    Sub RegHotKeys(hWnd As Long)
        RetHotKey1 = RegisterHotKey(hWnd, 0, MOD_CONTROL, vbKeyO) 'open
        RetHotKey2 = RegisterHotKey(hWnd, 1, MOD_CONTROL, vbKeyV) 'paste
        RetHotKey3 = RegisterHotKey(hWnd, 2, MOD_CONTROL, vbKeyS) 'save
        RetHotKey4 = RegisterHotKey(hWnd, 3, MOD_CONTROL, vbKeyX) 'hide
        RetHotKey5 = RegisterHotKey(hWnd, 4, MOD_CONTROL, vbKeyZ) 'show
        PrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf KeyCallbacks)
    End Sub
    Sub UnRegHotKeys(hWnd As Long)
        If RetHotKey1 Then
            UnregisterHotKey hWnd, 0
        End If
        If RetHotKey2 Then
            UnregisterHotKey hWnd, 1
        End If
        If RetHotKey3 Then
            UnregisterHotKey hWnd, 2
        End If
    End Sub
'==========================================================================================================
'Process HOT KEYS END
'==========================================================================================================

'==========================================================================================================
'Process System Tray
'==========================================================================================================

Sub tmrTray_Timer()
    icoIndex = icoIndex + 1
    CreateSysTrayIconAnim icoIndex
    If icoIndex > 12 Then
        icoIndex = 0
    End If
End Sub

'==========================================================================================================
'Process System Tray End
'==========================================================================================================

'==========================================================================================================
'Process System Menu Start
'==========================================================================================================

Sub MakeSystemMenu()
   Dim r As Long
   Dim hMenu As Long
   hMenu = GetSystemMenu(Me.hWnd, False)
   r = AppendMenu(hMenu, MF_SEPARATOR, 0, 0&)
   r = AppendMenu(hMenu, MF_STRING, ID_ABOUT, "&About..")
   r = AppendMenu(hMenu, MF_STRING, ID_TRAY, "&Minimized to Tray")
   If r = 1 Then
      Call HookWindow(Me.hWnd, Me)
   End If
End Sub

Friend Function WindowProc(hWnd As Long, msg As Long, wp As Long, lp As Long) As Long
   Select Case msg
      Case WM_SYSCOMMAND
         If wp = ID_ABOUT Then
            Form1.Show vbModal
            WindowProc = 1
            Exit Function
         End If
         If wp = ID_TRAY Then
            frmMain.Visible = False
            WindowProc = 1
            Exit Function
         End If
      Case Else
   End Select
   WindowProc = CallWindowProc(GetProp(hWnd, "OldWindowProc"), hWnd, msg, wp, lp)
End Function

Sub RestoreHook()
    Call UnhookWindow(Me.hWnd)
End Sub

'==========================================================================================================
'Process System Menu End
'==========================================================================================================

Private Sub tmrWait_Timer()
    SecCnt2 = SecCnt2 + 1
    If SecCnt2 = sWait15Min * 60 Then
       tmrWait.Enabled = False
       Call CmdDownload_Click
    End If
End Sub
