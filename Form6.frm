VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "程序选择"
   ClientHeight    =   3090
   ClientLeft      =   5370
   ClientTop       =   4155
   ClientWidth     =   4680
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "集装箱装配效率描绘"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "集装箱容量计算器"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      MaskColor       =   &H00FFC0C0&
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Form6.Hide

End Sub

Private Sub Command2_Click()
Form3.Show
Form6.Hide

End Sub
