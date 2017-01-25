VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Setting 
   Caption         =   "设置"
   ClientHeight    =   3036
   ClientLeft      =   4188
   ClientTop       =   2880
   ClientWidth     =   5268
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Setting"
   ScaleHeight     =   3036
   ScaleWidth      =   5268
   Begin VB.CommandButton Command5 
      Caption         =   "快捷键盘"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3840
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   1440
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      Caption         =   "是否显示"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Width           =   1332
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "颜色"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3480
      TabIndex        =   6
      Top             =   720
      Width           =   732
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2760
      TabIndex        =   5
      Top             =   2040
      Width           =   972
   End
   Begin VB.CommandButton Command2 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   972
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "Form2.frx":048A
      Left            =   1200
      List            =   "Form2.frx":048C
      TabIndex        =   3
      Text            =   "选择"
      Top             =   120
      Width           =   2652
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1200
      TabIndex        =   2
      Top             =   1440
      Width           =   2532
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4080
      TabIndex        =   0
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Height          =   372
      Left            =   3000
      TabIndex        =   7
      Top             =   720
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "y="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function getIndex(str As String) As Integer
Dim index As Integer
index = -1
Dim i As Integer
For i = 0 To Combo1.ListCount - 1
If Combo1.List(i) = str Then
index = i
Exit For
End If
Next
getIndex = index
End Function

Private Sub Combo1_Click()
Dim index As Integer
index = Data.getIndex(Combo1.Text)
If index <> -1 Then
Check1.Enabled = True
Command4.Enabled = True
Label3.BackColor = functionsColor(index)
Check1.Value = functionsEnable(index)
Else
Check1.Enabled = False
Command4.Enabled = False
End If
End Sub

Private Sub Command1_Click()
Unload Form3
Me.Hide
Main.Show
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" And Text1.Text <> " " Then
Call addFunctions(Text1.Text)
Combo1.AddItem (Text1.Text)
Combo1.Text = Text1.Text
Call Combo1_Click
Text1.Text = ""
End If
End Sub

Private Sub Command3_Click()
If getIndex(Combo1.Text) >= 0 Then
Call removeFunctions(Combo1.Text)
Combo1.RemoveItem (getIndex(Combo1.Text))
End If
If getIndex(Combo1.Text) = -1 Then
Check1.Enabled = False
Command4.Enabled = False
End If
End Sub

Private Sub Command4_Click()
CommonDialog1.ShowColor
Label3.BackColor = CommonDialog1.color
Call setColor(Combo1.Text, CommonDialog1.color)
End Sub

Private Sub check1_Click()
  If Check1.Value = 1 Then
  Call setEnable(Combo1.Text, 1)
  Else
  Call setEnable(Combo1.Text, 0)
  End If
End Sub

Private Sub Command5_Click()
Form3.Show
End Sub
