VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4425
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   0
      Picture         =   "Form1.frx":0ECA
      ScaleHeight     =   4245
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   4515
         Top             =   315
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   105
         Left            =   480
         TabIndex        =   2
         Top             =   3960
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   5250
         Top             =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "加载中请稍后....."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   480
         TabIndex        =   1
         Top             =   3675
         Width           =   2640
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'申皇装用
'****************************************************************************
Private loadform As Boolean
Private Sub Form_Load()
Timer1.Enabled = True
loadform = False
ProgressBar1.Value = 0
End Sub

Private Sub Timer1_Timer()
If Not loadform Then
    Load Form3
End If
ProgressBar1.Value = ProgressBar1.Value + 5
If ProgressBar1.Value = ProgressBar1.max Then
Form6.Show
Unload Me
End If
End Sub

Private Sub Timer2_Timer()
If Label1.Visible = False Then
Label1.Visible = True
Else
Label1.Visible = False
End If
End Sub
