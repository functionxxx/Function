VERSION 5.00
Begin VB.Form Main 
   Caption         =   "主界面"
   ClientHeight    =   4200
   ClientLeft      =   4008
   ClientTop       =   2316
   ClientWidth     =   7272
   LinkTopic       =   "Main"
   ScaleHeight     =   4200
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
      Top             =   2640
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
      Top             =   1440
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
      Top             =   240
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
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub setInitial()
X = Picture1.Width / 500
y = Picture1.Height / 500
Picture1.Scale (-X, y)-(X, -y)
Picture1.Line (-X, 0)-(X, 0), vbBlack
Picture1.Line (0, y)-(0, -y), vbBlack
Dim j
For j = -y To y
If j <> 0 Then
Call setPaintPosition(-1, j): Picture1.Print j
End If
Next j
Dim i
For i = -X To X
Call setPaintPosition(i, 0): Picture1.Print i
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
    For i = -X To X Step 0.01
    Dim y
    y = Replace(expression, "x", i)
    y = sctl.eval(y)
    Picture1.PSet (i, y), color
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
Call setInitial
End Sub

Private Sub Form1_Paint()
Call setInitial
End Sub

Private Sub setPaintPosition(ByVal X As Single, ByVal y As Single)
Picture1.CurrentX = X
Picture1.CurrentY = y
End Sub
