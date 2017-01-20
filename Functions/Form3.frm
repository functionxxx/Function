VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "输入键盘"
   ClientHeight    =   1932
   ClientLeft      =   11460
   ClientTop       =   2424
   ClientWidth     =   4224
   LinkTopic       =   "Form1"
   ScaleHeight     =   1932
   ScaleWidth      =   4224
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command28 
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
      Left            =   3000
      TabIndex        =   26
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command27 
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
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command26 
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
      Left            =   3600
      TabIndex        =   24
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command25 
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
      Left            =   3600
      TabIndex        =   23
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command23 
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
      Left            =   3000
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command22 
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
      Left            =   3000
      TabIndex        =   21
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command21 
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
      Left            =   3000
      TabIndex        =   20
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command20 
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
      Left            =   2400
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command19 
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
      Left            =   2400
      TabIndex        =   18
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command18 
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
      Left            =   2400
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command17 
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
      Left            =   2400
      TabIndex        =   16
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command16 
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
      Left            =   1800
      TabIndex        =   15
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command15 
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
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command14 
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
      Left            =   1800
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command13 
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
      Left            =   1800
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command12 
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command11 
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
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command10 
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
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command9 
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
      Left            =   1200
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command8 
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
      Left            =   600
      TabIndex        =   7
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command7 
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
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command6 
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
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command5 
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
      Left            =   600
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command4 
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
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Command3 
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
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command2 
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

Private Sub Command1_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "7"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "7"
End If
End Sub

Private Sub Command10_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "6"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "6"
End If
End Sub

Private Sub Command11_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "3"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "3"
End If
End Sub

Private Sub Command12_Click()
Setting.Text1.Text = Setting.Text1.Text + "-"
End Sub

Private Sub Command13_Click()
Setting.Text1.Text = Setting.Text1.Text + "+"
End Sub

Private Sub Command14_Click()
Setting.Text1.Text = Setting.Text1.Text + "-"
End Sub

Private Sub Command15_Click()
Setting.Text1.Text = Setting.Text1.Text + "*"
End Sub

Private Sub Command16_Click()
Setting.Text1.Text = Setting.Text1.Text + "/"
End Sub

Private Sub Command17_Click()
Setting.Text1.Text = Setting.Text1.Text + "^"
End Sub

Private Sub Command18_Click()
Setting.Text1.Text = Setting.Text1.Text + "x"
End Sub

Private Sub Command19_Click()
Setting.Text1.Text = Setting.Text1.Text + "sin("
End Sub

Private Sub Command2_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "4"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "4"
End If
End Sub

Private Sub Command20_Click()
Setting.Text1.Text = Setting.Text1.Text + "cos("
End Sub

Private Sub Command21_Click()
Setting.Text1.Text = Setting.Text1.Text + "tan("
End Sub

Private Sub Command22_Click()
Setting.Text1.Text = Setting.Text1.Text + "log("
End Sub

Private Sub Command23_Click()
Setting.Text1.Text = Setting.Text1.Text + "Abs("
End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub Command25_Click()
Setting.Text1.Text = Setting.Text1.Text + "("
End Sub

Private Sub Command26_Click()
Setting.Text1.Text = Setting.Text1.Text + ")"
End Sub

Private Sub Command27_Click()
Unload Form3
Setting.SetFocus
End Sub

Private Sub Command28_Click()
If Setting.Text1.Text <> "" Then
Setting.Text1.Text = Left(Setting.Text1.Text, Len(Setting.Text1.Text) - 1)
Else
Setting.Text1.Text = Setting.Text1.Text + ""
End If
End Sub
Private Sub Command3_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "1"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "1"
End If
End Sub

Private Sub Command4_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "0"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "0"
End If
End Sub

Private Sub Command5_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "8"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "8"
End If
End Sub

Private Sub Command6_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "5"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "5"
End If
End Sub

Private Sub Command7_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "2"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "2"
End If
End Sub

Private Sub Command8_Click()
If Setting.Text1.Text <> "" Then
Setting.Text1.Text = Setting.Text1.Text + "."
Else
Setting.Text1.Text = Setting.Text1.Text + "0."
End If
End Sub

Private Sub Command9_Click()
If Setting.Text1.Text <> "" Then
If Setting.Text1.Text = "0" Then
Setting.Text1.Text = "0"
Else
Setting.Text1.Text = Setting.Text1.Text + "9"
End If
Else
Setting.Text1.Text = Setting.Text1.Text + "9"
End If
End Sub


