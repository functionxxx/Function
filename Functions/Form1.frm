VERSION 5.00
Begin VB.Form Main 
   Caption         =   "主界面"
   ClientHeight    =   4308
   ClientLeft      =   4188
   ClientTop       =   2316
   ClientWidth     =   7272
   LinkTopic       =   "Main"
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   4308
   ScaleWidth      =   7272
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4800
      Top             =   3840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5400
      TabIndex        =   3
      Top             =   3240
      Width           =   1452
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5400
      TabIndex        =   2
      Top             =   1680
      Width           =   1452
   End
   Begin VB.CommandButton Command1 
      Caption         =   "绘画"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   1452
   End
   Begin VB.PictureBox Picture1 
      Height          =   3684
      Left            =   120
      ScaleHeight     =   3636
      ScaleWidth      =   4908
      TabIndex        =   0
      Top             =   120
      Width           =   4956
   End
   Begin VB.Label Label1 
      Caption         =   "X:     Y:"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   3012
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public oldx As Single
Public oldy As Single
Public zt As Integer

Private Sub setInitial(ByVal movex As Single, ByVal movey As Single)
Picture1.Cls
If movex = 0 And movey = 0 Then
Data.x1 = -(Picture1.Width / 840)
Data.y1 = (Picture1.Height / 840)
Data.x2 = (Picture1.Width / 840)
Data.y2 = -(Picture1.Height / 840)
Else
Data.x1 = Data.x1 + movex
Data.x2 = Data.x2 + movex
Data.y1 = Data.y1 + movey
Data.y2 = Data.y2 + movey
End If
Picture1.Scale (Data.x1, Data.y1)-(Data.x2, Data.y2)
Picture1.Line (Data.x1, 0)-(Data.x2, 0), vbBlack
Picture1.Line (0, Data.y1)-(0, Data.y2), vbBlack

Dim j
For j = Data.y2 To Data.y1
If j <> 0 Then
Call setPaintPosition(-1, j): Picture1.Print j
Picture1.DrawStyle = 2: Picture1.Line (Data.x1, j)-(Data.x2, j), vbbalck: Picture1.DrawStyle = 0
End If
Next j
Dim i
For i = Data.x1 To Data.x2
Call setPaintPosition(i, 0): Picture1.Print i
Picture1.DrawStyle = 2: Picture1.Line (i, Data.y1)-(i, Data.y2), vbbalck: Picture1.DrawStyle = 0
Next i
End Sub

Private Sub draw()
Dim sctl As Object
Set sctl = CreateObject("MSScriptControl.ScriptControl")
sctl.Language = "VBScript"

Dim i
Dim j
For j = 0 To (Data.count - 1) Step 1
  Dim expression As String
  Dim color As String
  expression = functions(j)
  color = functionsColor(j)
  If functionsEnable(j) = 1 Then
    For i = Data.x1 To Data.x2 Step 0.0005
    On Error Resume Next
    Picture1.PSet (i, sctl.eval(Replace(expression, "x", i))), color
    DoEvents
    Next
  End If
  DoEvents
Next

Set sctl = Nothing
End Sub

Private Sub Command1_Click()
If Data.count > 0 Then
Call draw
End If
End Sub

Private Sub Command2_Click()
Unload Me
Setting.Show
End Sub

Private Sub Command3_Click()
Unload Me
Unload Setting
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
Command1.Left = Main.Width - Command1.Width - 4 * 120
Command2.Left = Main.Width - Command2.Width - 4 * 120
Command3.Left = Main.Width - Command3.Width - 4 * 120
Picture1.Width = Command1.Left - 4 * 120
Picture1.Height = Main.Height - 10 * 120
Label1.Top = Picture1.Height + 120
Command3.Top = Label1.Top - Command3.Height
Command2.Top = (Command3.Top - Command1.Top) / 2

Call setInitial(0, 0)
End If
End Sub

Private Sub Form_Paint()
Call setInitial(0, 0)
End Sub

Private Sub setPaintPosition(ByVal xx As Single, ByVal yy As Single)
Picture1.CurrentX = xx
Picture1.CurrentY = yy
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oldx = X
  oldy = Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 0 Then
Dim sx, sy
sx = Format(X, ".00")
sy = Format(Y, ".00")
Label1.Caption = "X:" & IIf(sx > -1 And sx < 1, Format(sx, "0.00"), sx) & "  Y:" & IIf(sy > -1 And sy < 1, Format(sy, "0.00"), sy)
ElseIf Button = vbLeftButton Then
    Call setInitial(-(X - oldx), -(Y - oldy))
End If
End Sub
