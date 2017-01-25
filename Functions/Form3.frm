VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "输入键盘"
   ClientHeight    =   1932
   ClientLeft      =   9456
   ClientTop       =   2448
   ClientWidth     =   4224
   LinkTopic       =   "Form1"
   ScaleHeight     =   1932
   ScaleWidth      =   4224
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Bks ←"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   3000
      TabIndex        =   26
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter "
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   26
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   3600
      TabIndex        =   24
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   3600
      TabIndex        =   23
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Abs"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   3000
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "log"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   3000
      TabIndex        =   21
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   3000
      TabIndex        =   20
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   2400
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   2400
      TabIndex        =   18
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   2400
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   2400
      TabIndex        =   16
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   1800
      TabIndex        =   15
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1800
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   1800
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   1200
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Private oldx As Single
Private oldy As Single

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "7"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "7"
  End If
ElseIf Index = 1 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "4"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "4"
  End If
ElseIf Index = 2 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "1"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "1"
  End If
ElseIf Index = 3 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "0"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "0"
  End If
ElseIf Index = 4 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "8"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "8"
  End If
ElseIf Index = 5 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "5"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "5"
  End If
ElseIf Index = 6 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "2"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "2"
  End If
ElseIf Index = 7 Then
  If Setting.Text1.Text <> "" Then
  Setting.Text1.Text = Setting.Text1.Text + "."
  Else
  Setting.Text1.Text = Setting.Text1.Text + "0."
  End If
ElseIf Index = 8 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "9"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "9"
  End If
ElseIf Index = 9 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "6"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "6"
  End If
ElseIf Index = 10 Then
  If Setting.Text1.Text <> "" Then
  If Setting.Text1.Text = "0" Then
  Setting.Text1.Text = "0"
  Else
  Setting.Text1.Text = Setting.Text1.Text + "3"
  End If
  Else
  Setting.Text1.Text = Setting.Text1.Text + "3"
  End If
ElseIf Index = 11 Then
  Setting.Text1.Text = Setting.Text1.Text + "-"
ElseIf Index = 12 Then
  Setting.Text1.Text = Setting.Text1.Text + "+"
ElseIf Index = 13 Then
  Setting.Text1.Text = Setting.Text1.Text + "-"
ElseIf Index = 14 Then
  Setting.Text1.Text = Setting.Text1.Text + "*"
ElseIf Index = 15 Then
  Setting.Text1.Text = Setting.Text1.Text + "/"
ElseIf Index = 16 Then
  Setting.Text1.Text = Setting.Text1.Text + "^"
ElseIf Index = 17 Then
  Setting.Text1.Text = Setting.Text1.Text + "x"
ElseIf Index = 18 Then
  Setting.Text1.Text = Setting.Text1.Text + "sin("
ElseIf Index = 19 Then
  Setting.Text1.Text = Setting.Text1.Text + "cos("
ElseIf Index = 20 Then
  Setting.Text1.Text = Setting.Text1.Text + "tan("
ElseIf Index = 21 Then
  Setting.Text1.Text = Setting.Text1.Text + "log("
ElseIf Index = 22 Then
  Setting.Text1.Text = Setting.Text1.Text + "Abs("
ElseIf Index = 23 Then
  If Setting.Text1.Text <> "" Then
  Setting.Text1.Text = Left(Setting.Text1.Text, Len(Setting.Text1.Text) - 1)
  Else
  Setting.Text1.Text = Setting.Text1.Text + ""
  End If
ElseIf Index = 24 Then
  Setting.Text1.Text = Setting.Text1.Text + "("
ElseIf Index = 25 Then
  Setting.Text1.Text = Setting.Text1.Text + ")"
ElseIf Index = 26 Then
  Unload Form3
  Setting.SetFocus
End If
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   oldx = X
   oldy = Y
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    Form3.Move Form3.Left + X - oldx, Form3.Top + Y - oldy
    End If
End Sub
