VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "申皇集装箱计算器  Version1.0"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "楷体_GB2312"
      Size            =   26.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFF80&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10440
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "导入数据"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   27
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "结果显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5400
      TabIndex        =   10
      Top             =   4080
      Width           =   4695
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1095
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "集装箱规格"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5400
      TabIndex        =   3
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "米"
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "米"
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "米"
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "高度："
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "宽度："
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "长度："
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "货物规格"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "米"
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "米"
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "米"
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "高度："
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "宽度："
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "长度："
         BeginProperty Font 
            Name            =   "宋体-PUA"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重  置"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计  算"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   360
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Copyright to Hangzhou shenhuang Nonwovens Co.,Ltd                                                                    by:Authur Xu"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   0
      TabIndex        =   25
      Top             =   6480
      Width           =   10695
   End
   Begin VB.Image Image7 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":0ECA
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Image6 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":4212
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Image5 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":65C8
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Image4 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":89DA
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":AE76
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Image2 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":D31E
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   480
      Picture         =   "Form2.frx":F75C
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "长度："
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "摆放方式示意图例"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   1560
      TabIndex        =   26
      Top             =   4080
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim w As String
Dim l As String
Dim h As String
Dim X As String
Dim Y As String
Dim z As String
Dim a As Double
Dim b As Double
Dim c As Double
Dim num1 As Double
Dim num2 As Double
Dim num3 As Double
Dim num4 As Double
Dim num5 As Double
Dim num6 As Double
Dim u(5)
Dim max As Integer
Dim i As Integer
Dim k As Integer
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False

l = Text1
w = Text2
h = Text3
X = Text4
Y = Text5
z = Text6
num1 = Fix(Val(X) / Val(l)) * Fix(Val(Y) / Val(w)) * Fix(Val(z) / Val(h))
num2 = Fix(Val(X) / Val(w)) * Fix(Val(Y) / Val(l)) * Fix(Val(z) / Val(h))
num3 = Fix(Val(X) / Val(l)) * Fix(Val(Y) / Val(h)) * Fix(Val(z) / Val(w))
num4 = Fix(Val(X) / Val(h)) * Fix(Val(Y) / Val(l)) * Fix(Val(z) / Val(w))
num5 = Fix(Val(X) / Val(h)) * Fix(Val(Y) / Val(w)) * Fix(Val(z) / Val(l))
num6 = Fix(Val(X) / Val(w)) * Fix(Val(Y) / Val(h)) * Fix(Val(z) / Val(l))
u(0) = num1
u(1) = num2
u(2) = num3
u(3) = num4
u(4) = num5
u(5) = num6
max = u(0)
For i = 1 To 5
If max < u(i) Then
max = u(i)
k = i
End If
Next
If Val(l) = Val(w) And Val(w) = Val(h) Then
Image7.Visible = True
End If

If k = 1 Then
Image2.Visible = True
ElseIf k = 2 Then
Image3.Visible = True
ElseIf k = 3 Then
Image4.Visible = True
ElseIf k = 4 Then
Image5.Visible = True
ElseIf k = 5 Then
Image6.Visible = True
ElseIf max = num1 Then
Image1.Visible = True
End If

Text7 = max & "箱"









End Sub



Private Sub Picture1_Click()

End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""
Text6 = ""
Text7 = ""
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
Image5.Visible = False
Image6.Visible = False
Image7.Visible = False



End Sub

Private Sub Command3_Click()
Dim n As Integer
Dim m As Integer
Dim a As Double
Dim b As Double
Call Command1_Click
a = Val(InputBox("请输入每件货物的重量，单位为“千克”", "提示!!"))
If a = 0 Then
a = 1
End If
b = Val(InputBox("请输入容器装载重量，单位为“千克”", "提示!!"))
If b = 0 Then
b = 20000
End If
Form3.Show
Form3.Text1(1) = Val(Text1) * 1000
Form3.Text1(2) = Val(Text2) * 1000
Form3.Text1(3) = Val(Text3) * 1000
Form3.Text1(4) = a * 1000
Form3.Text1(5) = Val(Text7)
Form3.Text2(1) = Val(Text4) * 1000
Form3.Text2(2) = Val(Text5) * 1000
Form3.Text2(3) = Val(Text6) * 1000
Form3.Text2(4) = b * 1000

Form3.Text1(0) = "默认商品"
Form3.Text2(0) = "默认容器"


End Sub
