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

Private Sub setInitial()
Picture1.Cls
Module1.X = Picture1.Width / 600
Module1.Y = Picture1.Height / 600
Picture1.Scale (-Module1.X, Module1.Y)-(Module1.X, -Module1.Y)
Picture1.Line (-Module1.X, 0)-(Module1.X, 0), vbBlack
Picture1.Line (0, Module1.Y)-(0, -Module1.Y), vbBlack
Dim j
For j = -Module1.Y To Module1.Y
If j <> 0 Then
Call setPaintPosition(-1, j): Picture1.Print j
Picture1.DrawStyle = 2: Picture1.Line (-Module1.X, j)-(Module1.X, j), vbbalck: Picture1.DrawStyle = 0
End If
Next j
Dim i
For i = -Module1.X To Module1.X
Call setPaintPosition(i, 0): Picture1.Print i
Picture1.DrawStyle = 2: Picture1.Line (i, Module1.Y)-(i, -Module1.Y), vbbalck: Picture1.DrawStyle = 0
Next i
End Sub

Private Sub draw()
Dim sctl As Object
Set sctl = CreateObject("MSScriptControl.ScriptControl")
sctl.Language = "VBScript"

Dim i
Dim j
For j = 0 To (Module1.count - 1) Step 1
  Dim expression As String
  Dim color As String
  expression = functions(j)
  color = functionsColor(j)
  If functionsEnable(j) Then
    For i = -Module1.X To Module1.X Step 0.005
    Dim Y
    Y = Replace(expression, "x", i)
    On Error Resume Next
    Y = sctl.eval(Y)
    Picture1.PSet (i, Y), color
    Next i
  End If
Next j
End Sub

Private Sub Command1_Click()
If Setting.Combo1.ListIndex <> -1 Then
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
Command1.Left = Main.Width - Command1.Width - 4 * 120
Command2.Left = Main.Width - Command2.Width - 4 * 120
Command3.Left = Main.Width - Command3.Width - 4 * 120
Picture1.Width = Command1.Left - 4 * 120
Picture1.Height = Main.Height - 10 * 120
Label1.Top = Picture1.Height + 120
Command3.Top = Label1.Top - Command3.Height
Command2.Top = (Command3.Top - Command1.Top) / 2

Call setInitial
End Sub

Private Sub Form1_Paint()
Call setInitial
End Sub

Private Sub setPaintPosition(ByVal xx As Single, ByVal yy As Single)
Picture1.CurrentX = xx
Picture1.CurrentY = yy
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim sx, sy
sx = Format(X, ".00")
sy = Format(Y, ".00")
Label1.Caption = "X:" & IIf(sx > -1 And sx < 1, Format(sx, "0.00"), sx) & "  Y:" & IIf(sy > -1 And sy < 1, Format(sy, "0.00"), sy)
End Sub
