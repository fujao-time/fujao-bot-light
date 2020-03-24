VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Fujao-Bot Light [1.000]"
   ClientHeight    =   6735
   ClientLeft      =   540
   ClientTop       =   720
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9375
   Begin VB.CommandButton Command2 
      Caption         =   "Status: OFF"
      Height          =   615
      Left            =   4800
      TabIndex        =   7
      Top             =   6000
      Width           =   4455
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2520
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Interval        =   60000
      Left            =   2040
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gacha!"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   8640
      TabIndex        =   5
      Top             =   120
      Width           =   645
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   7920
      TabIndex        =   4
      Top             =   120
      Width           =   645
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   645
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1560
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Microsoft YaHei UI"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6420
      Left            =   9480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":143E1
      Top             =   120
      Width           =   7005
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5310
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   9366
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ready As Boolean
Dim pwr As Boolean
Dim the_void As Long
Dim slist As Variant
Dim scount As Long

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

Private Sub Command1_Click()
MsgBox Text1.Text
End Sub

Private Sub Command2_Click()
If pwr Then
pwr = False
Command2.Caption = "Status: OFF"
Else
pwr = True
Command2.Caption = "Status: ON"
End If
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "about:blank"
Text2.Text = "https://cdn.jsdelivr.net/gh/fujao-time/fujaoese-hitokoto/sentence.txt"
Text4.Text = 0
Text5.Text = 0
ready = True
pwr = False
End Sub

Private Sub Timer1_Timer() 'parse
If WebBrowser1.busy Then
Exit Sub
Else
slist = Split(WebBrowser1.Document.body.innerText, vbCrLf)
scount = Fix(Val(slist(0)))
Timer1.Enabled = False
Timer4.Enabled = True
End If
End Sub

Private Sub Timer2_Timer() 'fetch sentences
Timer4.Enabled = False
WebBrowser1.Navigate Text2.Text
Timer1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer() 'check time
Text3.Text = Minute(Now)
If pwr Then
If Val(Text3.Text) Mod 30 = 0 And ready Then
Timer6.Enabled = True
ready = False
ElseIf Val(Text3.Text) Mod 30 <> 0 Then
ready = True
End If
End If
End Sub

Private Sub Timer4_Timer() 'update text
Dim x As Long
x = Fix(Rnd * scount) + 1
If x >= 1 And x <= scount Then
Text1.Text = "[Fujao-time]" & vbCrLf & Year(Now) & "/" & Month(Now) & "/" & Day(Now) & " " & Hour(Now) & ":" & Text3.Text & vbCrLf & slist(x)
End If
End Sub

Private Sub Timer5_Timer() 'check update every 10 min
Timer2.Enabled = True
End Sub

Private Sub Timer6_Timer() 'send message
Clipboard.Clear
Clipboard.SetText Text1.Text
the_void = SetCursorPos(Val(Text4.Text), Val(Text5.Text))
mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
SendKeys ("^V")
SendKeys ("^{ENTER}")
Timer6.Enabled = False
End Sub
