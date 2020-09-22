VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4065
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Rapidshare Downloader v.2.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   1
         Top             =   540
         Width           =   3585
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Me.Width = 4455
    Me.Height = 2730
    Me.Icon = frmMain.Icon
End Sub
