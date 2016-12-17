VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Setting 
   Caption         =   "设置"
   ClientHeight    =   3036
   ClientLeft      =   5676
   ClientTop       =   2880
   ClientWidth     =   5268
   LinkTopic       =   "Setting"
   ScaleHeight     =   3036
   ScaleWidth      =   5268
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
      ItemData        =   "Form2.frx":0000
      Left            =   1200
      List            =   "Form2.frx":0002
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
      Text            =   "0"
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

Private Sub Combo1_Click()
Dim index As Integer
index = getIndex(Combo1.Text)
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
Me.Hide
Main.Show
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" And Text1.Text <> " " Then
Call addFunctions(Text1.Text)
Combo1.AddItem (Text1.Text)
Text1.Text = ""
End If
End Sub

Private Sub Command3_Click()
If Combo1.ListIndex >= 0 Then
Call removeFunctions(Combo1.Text)
Combo1.RemoveItem (Combo1.ListIndex)
End If
If Combo1.ListIndex = -1 Then
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
