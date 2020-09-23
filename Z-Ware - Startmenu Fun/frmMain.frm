VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Z-Ware - Start Menu Fun                     "
   ClientHeight    =   1035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1035
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer trmScroller 
      Interval        =   100
      Left            =   960
      Top             =   1440
   End
   Begin VB.Frame fraStartCaption 
      Caption         =   "Start Button"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdShowStart 
         Caption         =   "Show"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdHideStart 
         Caption         =   "Hide"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtStartCaption 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStartCaption 
         Caption         =   "Caption:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   260
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdHideStart_Click()
Call ShowWindow(Info.StartButton, SW_HIDE)
End Sub

Private Sub cmdShowStart_Click()
Call ShowWindow(Info.StartButton, SW_SHOW)
End Sub

Private Sub Form_Load()
x = FindWindow("Shell_TrayWnd", vbNullString)
Info.StartButton = FindWindowEx(x, 0&, "Button", vbNullString)
End Sub

Private Sub trmScroller_Timer()
Dim sTitle As String
sTitle = Me.Caption
sTitle = Mid(sTitle, 2, Len(sTitle)) & Mid(sTitle, 1, 1)
Me.Caption = sTitle
End Sub

Private Sub txtStartCaption_Change()
Call SendMessageString(Info.StartButton, WM_SETTEXT, 0, txtStartCaption)
End Sub
