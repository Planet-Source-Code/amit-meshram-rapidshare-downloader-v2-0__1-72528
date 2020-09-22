VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUploader 
   Caption         =   "Rapidshare Uploader..."
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6810
   LinkTopic       =   "Form2"
   ScaleHeight     =   2700
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   2280
      Width           =   1035
   End
   Begin VB.CommandButton CmdUpload 
      Caption         =   "Upload"
      Height          =   375
      Left            =   3990
      TabIndex        =   11
      Top             =   2280
      Width           =   1035
   End
   Begin InetCtlsObjects.Inet Uploader 
      Left            =   150
      Top             =   2610
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "File Information"
      Height          =   1725
      Left            =   120
      TabIndex        =   3
      Top             =   510
      Width           =   6555
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   180
         TabIndex        =   8
         Top             =   1350
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label7 
         Caption         =   "@@@"
         Height          =   195
         Left            =   1890
         TabIndex        =   10
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Uploading Speed"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "@@@"
         Height          =   195
         Left            =   1890
         TabIndex        =   7
         Top             =   630
         Width           =   4005
      End
      Begin VB.Label Label4 
         Caption         =   "File Length"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "@@@"
         Height          =   195
         Left            =   1890
         TabIndex        =   5
         Top             =   360
         Width           =   4005
      End
      Begin VB.Label Label2 
         Caption         =   "File Name"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   1065
      End
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "Browse"
      Height          =   285
      Left            =   5850
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1395
   End
End
Attribute VB_Name = "frmUploader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
