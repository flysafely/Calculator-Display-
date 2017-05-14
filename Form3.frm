VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "申皇装箱效率计算程序"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   615
   ClientWidth     =   13770
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13770
   Begin VB.Frame Frame7 
      Caption         =   "装箱容器选择"
      Height          =   9045
      Left            =   105
      TabIndex        =   22
      Top             =   150
      Width           =   6840
      Begin VB.OptionButton Option1 
         Caption         =   "使用托盘(若要取出托盘,请不要选择)"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   56
         Top             =   4620
         Width           =   3675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "使用集装箱"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   55
         Top             =   210
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "常见托盘尺寸表"
         ForeColor       =   &H80000008&
         Height          =   3525
         Left            =   105
         TabIndex        =   35
         Top             =   5415
         Width           =   6615
         Begin VB.CommandButton Command11 
            Caption         =   "增加"
            Height          =   300
            Left            =   2835
            TabIndex        =   54
            Top             =   2880
            Width           =   1000
         End
         Begin VB.CommandButton Command10 
            Caption         =   "修改"
            Height          =   300
            Left            =   4095
            TabIndex        =   53
            Top             =   2880
            Width           =   1000
         End
         Begin VB.CommandButton Command9 
            Caption         =   "删除"
            Height          =   300
            Left            =   5355
            TabIndex        =   52
            Top             =   2880
            Width           =   1000
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   0
            Left            =   840
            TabIndex        =   41
            Top             =   2220
            Width           =   1100
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   1
            Left            =   3045
            TabIndex        =   40
            Top             =   2220
            Width           =   1100
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   2
            Left            =   5250
            TabIndex        =   39
            Top             =   2220
            Width           =   1100
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   3
            Left            =   840
            TabIndex        =   38
            Top             =   2535
            Width           =   1100
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   4
            Left            =   3045
            TabIndex        =   37
            Top             =   2535
            Width           =   1100
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   5
            Left            =   5250
            TabIndex        =   36
            Top             =   2535
            Width           =   1100
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1680
            Left            =   120
            TabIndex        =   42
            Top             =   315
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2963
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "选择"
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "类型"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "长(mm)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "宽(mm)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "高(mm)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "载重(g)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "自高(mm)"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "名称"
            Height          =   255
            Index           =   6
            Left            =   210
            TabIndex        =   48
            Top             =   2220
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "长(mm)"
            Height          =   255
            Index           =   12
            Left            =   2310
            TabIndex        =   47
            Top             =   2220
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "宽(mm)"
            Height          =   255
            Index           =   13
            Left            =   4515
            TabIndex        =   46
            Top             =   2220
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "载重(g)"
            Height          =   255
            Index           =   14
            Left            =   2310
            TabIndex        =   45
            Top             =   2535
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "高(mm)"
            Height          =   255
            Index           =   15
            Left            =   210
            TabIndex        =   44
            Top             =   2535
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "自高(mm)"
            Height          =   255
            Index           =   16
            Left            =   4515
            TabIndex        =   43
            Top             =   2535
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "常见集装箱尺寸表(多个集装箱可多选)"
         ForeColor       =   &H80000008&
         Height          =   4000
         Left            =   105
         TabIndex        =   23
         Top             =   525
         Width           =   6615
         Begin VB.CommandButton Command8 
            Caption         =   "增加"
            Height          =   300
            Left            =   2835
            TabIndex        =   51
            Top             =   3600
            Width           =   1000
         End
         Begin VB.CommandButton Command7 
            Caption         =   "修改"
            Height          =   300
            Left            =   4095
            TabIndex        =   50
            Top             =   3600
            Width           =   1000
         End
         Begin VB.CommandButton Command6 
            Caption         =   "删除"
            Height          =   300
            Left            =   5355
            TabIndex        =   49
            Top             =   3600
            Width           =   1000
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   0
            Left            =   840
            TabIndex        =   28
            Top             =   2940
            Width           =   1100
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   1
            Left            =   3045
            TabIndex        =   27
            Top             =   2940
            Width           =   1100
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   2
            Left            =   5250
            TabIndex        =   26
            Top             =   2940
            Width           =   1100
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   3
            Left            =   840
            TabIndex        =   25
            Top             =   3255
            Width           =   1100
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   270
            Index           =   4
            Left            =   3045
            TabIndex        =   24
            Top             =   3255
            Width           =   1100
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   2400
            Left            =   105
            TabIndex        =   29
            Top             =   315
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   4233
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "选择"
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "类型"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "长(mm)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "宽(mm)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "高(mm)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "载重(g)"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "高(mm)"
            Height          =   255
            Index           =   7
            Left            =   210
            TabIndex        =   34
            Top             =   3255
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "载重(g)"
            Height          =   255
            Index           =   8
            Left            =   2310
            TabIndex        =   33
            Top             =   3255
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "宽(mm)"
            Height          =   255
            Index           =   9
            Left            =   4515
            TabIndex        =   32
            Top             =   2940
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "长(mm)"
            Height          =   255
            Index           =   10
            Left            =   2310
            TabIndex        =   31
            Top             =   2940
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "名称"
            Height          =   255
            Index           =   11
            Left            =   210
            TabIndex        =   30
            Top             =   2940
            Width           =   975
         End
      End
      Begin VB.Label Label2 
         Caption         =   "说明：该选项用于计算托盘摆放方式，用于节省托盘使用情况，如果进仓不需要托盘，请不要选择该项目！"
         Height          =   735
         Left            =   480
         TabIndex        =   68
         Top             =   4920
         Width           =   4335
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "计算设置选项"
      Height          =   3855
      Left            =   7200
      TabIndex        =   1
      Top             =   5280
      Width           =   6495
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         Caption         =   "装箱货品尺寸种类"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   210
         TabIndex        =   69
         Top             =   240
         Width           =   6105
         Begin VB.OptionButton Option4 
            Caption         =   "多尺寸货品"
            Height          =   255
            Left            =   4080
            TabIndex        =   71
            Top             =   170
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            Caption         =   "单一尺寸货品"
            Height          =   255
            Left            =   1320
            TabIndex        =   70
            Top             =   170
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "返  回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         Caption         =   "剩余空间搜索策略"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   210
         TabIndex        =   63
         Top             =   2280
         Width           =   6105
         Begin VB.OptionButton Option2 
            Caption         =   "浅度搜索"
            Height          =   225
            Index           =   1
            Left            =   4215
            TabIndex        =   65
            Top             =   195
            Width           =   1485
         End
         Begin VB.OptionButton Option2 
            Caption         =   "深度搜索"
            Height          =   225
            Index           =   0
            Left            =   1275
            TabIndex        =   64
            Top             =   195
            Value           =   -1  'True
            Width           =   1485
         End
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "重  置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3000
         Width           =   1365
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "计  算"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "工作面拆分策略"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   210
         TabIndex        =   19
         Top             =   1560
         Width           =   6105
         Begin VB.CheckBox Check2 
            Caption         =   "浅度搜索"
            Height          =   225
            Index           =   1
            Left            =   4200
            TabIndex        =   62
            Top             =   195
            Value           =   1  'Checked
            Width           =   1650
         End
         Begin VB.CheckBox Check2 
            Caption         =   "深度搜索"
            Height          =   225
            Index           =   0
            Left            =   1320
            TabIndex        =   61
            Top             =   195
            Value           =   1  'Checked
            Width           =   2010
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "装箱优先策略"
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   210
         TabIndex        =   18
         Top             =   840
         Width           =   6105
         Begin VB.CheckBox Check1 
            Caption         =   "体积优先"
            Height          =   225
            Index           =   3
            Left            =   4725
            TabIndex        =   60
            Top             =   195
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "数量优先"
            Height          =   225
            Index           =   2
            Left            =   3255
            TabIndex        =   59
            Top             =   195
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "长大优先"
            Height          =   225
            Index           =   1
            Left            =   1800
            TabIndex        =   58
            Top             =   195
            Value           =   1  'Checked
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "宽大优先"
            Height          =   225
            Index           =   0
            Left            =   315
            TabIndex        =   57
            Top             =   195
            Value           =   1  'Checked
            Width           =   1275
         End
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   5640
         Picture         =   "Form3.frx":0ECA
         Top             =   3000
         Width           =   720
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "待装箱货物列表(混装货物可多选)"
      Height          =   5115
      Left            =   7140
      TabIndex        =   0
      Top             =   150
      Width           =   6495
      Begin VB.CommandButton Command13 
         Caption         =   "调用计算器"
         Height          =   300
         Left            =   1320
         TabIndex        =   67
         Top             =   4560
         Width           =   1245
      End
      Begin VB.CommandButton Command3 
         Caption         =   "删除"
         Height          =   300
         Left            =   5280
         TabIndex        =   17
         Top             =   4560
         Width           =   1000
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   300
         Left            =   4080
         TabIndex        =   16
         Top             =   4560
         Width           =   1000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "增加"
         Height          =   300
         Left            =   2880
         TabIndex        =   15
         Top             =   4560
         Width           =   1000
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   5
         Left            =   1320
         TabIndex        =   13
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   4200
         TabIndex        =   11
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   4
         Left            =   1320
         TabIndex        =   9
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   4200
         TabIndex        =   7
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   4200
         TabIndex        =   5
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   3480
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3000
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "选择"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "名称"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "长(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "宽(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "高(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "重量(g)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "数量"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "数量"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   14
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "高(mm)"
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   12
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "重量(g)"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   10
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "宽(mm)"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "长(mm)"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   6
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "名称"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   3480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'申皇专用
'****************************************************************************
'==========================================货物列表编辑==================================================
'增加货物列表
Private Sub Command1_Click()
        Set itmx = ListView3.ListItems.Add(, , "")
        For i = 1 To 6
            itmx.SubItems(i) = Text1(i - 1).Text
        Next i
        itmx.SubItems(6) = CStr(CInt(Text1(5).Text))
        Open App.Path & "\goods.txt" For Append As #1
Print #1, Text1(0).Text; "|"; Text1(1).Text; "|"; Text1(2).Text; "|"; Text1(3).Text; "|"; Text1(4).Text; "|"; Text1(5).Text; vbCrLf;
Close #1
        
End Sub

Private Sub Command12_Click()
Form6.Show

End Sub

Private Sub Command13_Click()
Form2.Show
End Sub

'修改货物列表
Private Sub Command2_Click()
Open App.Path & "\goods.txt" For Append As #1
Print #1, Text1(0).Text; "|"; Text1(1).Text; "|"; Text1(2).Text; "|"; Text1(3).Text; "|"; Text1(4).Text; "|"; Text1(5).Text; vbCrLf;
Close #1
ListView3.SelectedItem.SubItems(1) = Text1(0).Text
ListView3.SelectedItem.SubItems(2) = Text1(1).Text
ListView3.SelectedItem.SubItems(3) = Text1(2).Text
ListView3.SelectedItem.SubItems(4) = Text1(3).Text
ListView3.SelectedItem.SubItems(5) = Text1(4).Text
ListView3.SelectedItem.SubItems(6) = CStr(CInt(Text1(5).Text))
End Sub
'删除货物列表
Private Sub Command3_Click()
If ListView3.SelectedItem.Text = "" Then
MsgBox "请选择要删除的行！"
ElseIf ListView3.SelectedItem.Text = " ★" Then
a = MsgBox("确定要从列表中删除名称为" & ListView3.SelectedItem.SubItems(1) & "的货物数据么？", 308, "删除确认！")
If a = 6 Then
ListView3.ListItems.Remove (ListView3.SelectedItem.Index)
End If
End If
End Sub
'========================================================================================================


'=========================================集装箱列表编辑=================================================
'增加集装箱列表
Private Sub Command8_Click()
        Set itmx = ListView1.ListItems.Add(, , "")
        For i = 1 To 5
            itmx.SubItems(i) = Text2(i - 1).Text
        Next i
        Open App.Path & "\containers.txt" For Append As #1
Print #1, Text2(0).Text; "|"; Text2(1).Text; "|"; Text2(2).Text; "|"; Text2(3).Text; "|"; Text2(4).Text; vbCrLf;
Close #1
End Sub
'修改集装箱列表
Private Sub Command7_Click()
Open App.Path & "\containers.txt" For Append As #1
Print #1, Text2(0).Text; "|"; Text2(1).Text; "|"; Text2(2).Text; "|"; Text2(3).Text; "|"; Text2(4).Text; vbCrLf;
Close #1
ListView1.SelectedItem.SubItems(1) = Text2(0).Text
ListView1.SelectedItem.SubItems(2) = Text2(1).Text
ListView1.SelectedItem.SubItems(3) = Text2(2).Text
ListView1.SelectedItem.SubItems(4) = Text2(3).Text
ListView1.SelectedItem.SubItems(5) = Text2(4).Text

End Sub
'删除集装箱列表
Private Sub Command6_Click()
If ListView1.SelectedItem.Text = "" Then
MsgBox "请选择要删除的行！"
ElseIf ListView1.SelectedItem.Text = " ★" Then
a = MsgBox("确定要从列表中删除名称为" & ListView1.SelectedItem.SubItems(1) & "的集装箱数据么？", 308, "删除确认！")
If a = 6 Then
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End If
End If
End Sub
'========================================================================================================

'==========================================托盘列表编辑==================================================
'增加托盘列表
Private Sub Command11_Click()
        Set itmx = ListView2.ListItems.Add(, , "")
        For i = 1 To 6
            itmx.SubItems(i) = Text1(i - 1).Text
        Next i
        Open App.Path & "\trays.txt" For Append As #1
Print #1, Text3(0).Text; "|"; Text3(1).Text; "|"; Text3(2).Text; "|"; Text3(3).Text; "|"; Text3(4).Text; "|"; Text3(5).Text; vbCrLf;
Close #1
End Sub
'修改托盘列表
Private Sub Command10_Click()
Open App.Path & "\trays.txt" For Append As #1
Print #1, Text3(0).Text; "|"; Text3(1).Text; "|"; Text3(2).Text; "|"; Text3(3).Text; "|"; Text3(4).Text; "|"; Text3(5).Text; vbCrLf;
Close #1
ListView2.SelectedItem.SubItems(1) = Text3(0).Text
ListView2.SelectedItem.SubItems(2) = Text3(1).Text
ListView2.SelectedItem.SubItems(3) = Text3(2).Text
ListView2.SelectedItem.SubItems(4) = Text3(3).Text
ListView2.SelectedItem.SubItems(5) = Text3(4).Text
ListView2.SelectedItem.SubItems(6) = Text3(5).Text
End Sub
'删除托盘列表
Private Sub Command9_Click()
If ListView2.SelectedItem.Text = "" Then
MsgBox "请选择要删除的行！"
ElseIf ListView2.SelectedItem.Text = " ★" Then
a = MsgBox("确定要从列表中删除名称为" & ListView2.SelectedItem.SubItems(1) & "的托盘数据么？", 308, "删除确认！")
If a = 6 Then
ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End If
End If
End Sub
'========================================================================================================

'=================================判断输入是否为数字=====================================================
Private Function checkinput(obj As TextBox, num_type As Integer) As Boolean
If IsNumeric(obj.Text) Then
    Select Case num_type
    Case 0 '只要是数字
        checkinput = True
    Case 1 '必须为整数
        If CSng(obj.Text) Mod 1 = 0 Then
            checkinput = True
        Else
            checkinput = False
        End If
    Case Else
        checkinput = False
    End Select
Else
    checkinput = False
End If
End Function
Private Sub Command4_Click()
    '判断使用集装箱还是托盘并判断列表中是否有选中的集装箱或者托盘
    If Option1(0).Value = True Then
        Dim containers As Boolean
        containers = False
        For Each items In ListView1.ListItems
            If items.Text = " ★" Then containers = True
        Next
        If Not containers Then
            MsgBox "没有选择集装箱", 48, "错误！"
            Exit Sub
        End If
    ElseIf Option1(1).Value = True Then
        Dim trays As Boolean
        trays = False
        For Each items In ListView2.ListItems
            If items.Text = " ★" Then trays = True
        Next
        If Not trays Then
            MsgBox "没有选择托盘", 48, "错误！"
            Exit Sub
        End If
    End If
    
    '检查是否有待装货物
    Dim goods As Boolean
    goods = False
    For Each items In ListView3.ListItems
        If items.Text = " ★" And CInt(items.SubItems(6)) > 0 Then goods = True
    Next
    If Not goods Then
        MsgBox "没有选择要装箱的货物或者要装箱的货物数量为0", 48, "错误！"
        Exit Sub
    End If
    '检查装箱策略选择
    Dim check1flag As Boolean
    check1flag = False
    For Each check In Check1
        If check.Value = 1 Then check1flag = True
    Next
    If Not check1flag Then
        MsgBox "没有选择装箱策略", 48, "错误！"
        Exit Sub
    End If
    '检查工作面拆分策略选择
    Dim check2flag As Boolean
    check2flag = False
    For Each check In Check2
        If check.Value = 1 Then check2flag = True
    Next
    If Not check2flag Then
        MsgBox "没有选择工作面拆分策略", 48, "错误！"
        Exit Sub
    End If
    '检查剩余空间拆分策略选择
    Dim check3flag As Boolean
    check3flag = False
    For Each check In Option2
        If check.Value = True Then check3flag = True
    Next
    If Not check3flag Then
        MsgBox "没有选择剩余空间拆分策略", 48, "错误！"
        Exit Sub
    End If
    '验证通过，加载清单窗体
    Load Form4
    '选择的容器
    If Option1(0).Value = True Then
        For Each items In ListView1.ListItems
            If items.Text = " ★" Then
                Set itmx = Form4.ListView2.ListItems.Add(, , " ★")
                For i = 1 To 5
                    itmx.SubItems(i) = items.SubItems(i)
                Next i
            End If
        Next
        hc = True 'New Code
    ElseIf Option1(1).Value = True Then
        For Each items In ListView2.ListItems
            If items.Text = " ★" Then
                Set itmx = Form4.ListView2.ListItems.Add(, , " ★")
                For i = 1 To 6
                    itmx.SubItems(i) = items.SubItems(i)
                Next i
            End If
        Next
        hc = False 'New Code
    End If
    '选择的货物
    For Each items In ListView3.ListItems
            If items.Text = " ★" And CInt(items.SubItems(6)) > 0 Then
                Set itmx = Form4.ListView3.ListItems.Add(, , " ★")
                For i = 1 To 6
                    itmx.SubItems(i) = items.SubItems(i)
                Next i
            End If
    Next
    '选择的优先策略
    For Each check In Check1
        If check.Value = 1 Then
            Form4.Check1(check.Index).Value = 1
        Else
            Form4.Check1(check.Index).Value = 0
        End If
    Next
    '选择的拆分策略
    For Each check In Check2
        If check.Value = 1 Then
            Form4.Check2(check.Index).Value = 1
        Else
            Form4.Check2(check.Index).Value = 0
        End If
    Next
    '剩余空间拆分策略
    For Each check In Option2
        If check.Value = True Then
            Form4.Option2(check.Index).Value = True
        Else
            Form4.Option2(check.Index).Value = False
        End If
    Next
    Form4.getcount
    Form4.Show
End Sub
'========================================================================================================
Private Sub Command5_Click()
For Each items In ListView1.ListItems
    If items.Text = " ★" Then MsgBox "选中了" + CStr(items.Index)
Next
For Each items In ListView1.ListItems
    items.Text = ""
Next
For Each items In ListView2.ListItems
    items.Text = ""
Next
For Each items In ListView3.ListItems
    items.Text = ""
Next
Text2(0) = ""
Text2(1) = ""
Text2(2) = ""
Text2(3) = ""
Text2(4) = ""

Text1(0) = ""
Text1(1) = ""
Text1(2) = ""
Text1(3) = ""
Text1(4) = ""
Text1(5) = ""

Text3(0) = ""
Text3(1) = ""
Text3(2) = ""
Text3(3) = ""
Text3(4) = ""
Text3(5) = ""

End Sub
'========================================================================================================

Private Sub Form_Load()
initlistview '初始化listview里的数据
hc = False 'New Code
End Sub
'========================================================================================================

'初始化listview里的数据
Public Sub initlistview()
'集装箱列表
Open App.Path + "\containers.txt" For Input As #1
Seek #1, 1
ListView1.ListItems.Clear
i = 0
    Do While Not EOF(1)   ' 循环至文件尾。
        i = i + 1
        
        Line Input #1, textline   ' 读入一行数据。
        temps = Split(textline, "|")



        Set itmx = ListView1.ListItems.Add(, , "")
        For i = 0 To UBound(temps)
            itmx.SubItems(i + 1) = temps(i)
        Next i
        DoEvents
    Loop

Close
'托盘列表
Open App.Path + "\trays.txt" For Input As #1
Seek #1, 1
ListView2.ListItems.Clear
i = 0
    Do While Not EOF(1)   ' 循环至文件尾。
        i = i + 1
        
        Line Input #1, textline   ' 读入一行数据。
        temps = Split(textline, "|")



        Set itmx = ListView2.ListItems.Add(, , "")
        For i = 0 To UBound(temps)
            itmx.SubItems(i + 1) = temps(i)
        Next i
        DoEvents
    Loop

Close
'货物列表
Open App.Path + "\goods.txt" For Input As #1
Seek #1, 1
ListView3.ListItems.Clear
i = 0
    Do While Not EOF(1)   ' 循环至文件尾。
        i = i + 1
        
        Line Input #1, textline   ' 读入一行数据。
        temps = Split(textline, "|")



        Set itmx = ListView3.ListItems.Add(, , "")
        For i = 0 To UBound(temps)
            itmx.SubItems(i + 1) = temps(i)
        Next i
        If itmx.SubItems(6) = "" Then itmx.SubItems(6) = "0"
        
        If CInt(itmx.SubItems(6)) > 0 Then itmx.Text = " "
        DoEvents
    Loop

Close
End Sub
'========================================================================================================

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Text = "" Then
Item.Text = " ★"
Else
Item.Text = ""
End If
Text2(0).Text = Item.SubItems(1)
Text2(1).Text = Item.SubItems(2)
Text2(2).Text = Item.SubItems(3)
Text2(3).Text = Item.SubItems(4)
Text2(4).Text = Item.SubItems(5)
End Sub

Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Text = "" Then
Item.Text = " ★"
Else
Item.Text = ""
End If
Text3(0).Text = Item.SubItems(1)
Text3(1).Text = Item.SubItems(2)
Text3(2).Text = Item.SubItems(3)
Text3(3).Text = Item.SubItems(4)
Text3(4).Text = Item.SubItems(5)
Text3(5).Text = Item.SubItems(6)
End Sub

Private Sub ListView3_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Item.Text = "" Then
Item.Text = " ★"
Else
Item.Text = ""
End If

Text1(0).Text = Item.SubItems(1)
Text1(1).Text = Item.SubItems(2) 'x
Text1(2).Text = Item.SubItems(3) 'Y
Text1(3).Text = Item.SubItems(4) 'Z
Text1(4).Text = Item.SubItems(5)
Text1(5).Text = Item.SubItems(6)

End Sub


Private Sub Option3_Click()
If Option3.Value = True Then
Check1(0).Value = False
Check1(1).Value = False
Check1(3).Value = False
Check2(0).Value = False
Option2(0).Value = True
Option2(1).Value = False
Check1(0).Visible = False
Check1(1).Visible = False
Check1(3).Visible = False
Check2(0).Visible = False
Option2(0).Visible = True
Option2(1).Visible = False
End If
End Sub

Private Sub Option4_Click()
Check1(0).Visible = True
Check1(1).Visible = True
Check1(3).Visible = True
Check2(0).Visible = True
Option2(0).Visible = True
Option2(1).Visible = True
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index > 0 And Index < 5 Then
    If KeyAscii > 57 Or KeyAscii < 1 Or KeyAscii = 47 Then
        MsgBox "输入格式错误，请检查！当前要求输入数字。", 48, "输入错误！"
        KeyAscii = 0
    End If
ElseIf Index = 5 Then
    If KeyAscii > 57 Or KeyAscii < 1 Or KeyAscii = 47 Or KeyAscii = 96 Or KeyAscii = 48 Then
        MsgBox "输入格式错误，请检查！当前要求输入整数。", 48, "输入错误！"
        KeyAscii = 0
    End If
End If

End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If Index > 0 Then
    If KeyAscii > 57 Or KeyAscii < 1 Or KeyAscii = 47 Then
        MsgBox "输入格式错误，请检查！当前要求输入数字。", 48, "输入错误！"
        KeyAscii = 0
    End If
End If
End Sub
Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
If Index > 0 Then
    If KeyAscii > 57 Or KeyAscii < 1 Or KeyAscii = 47 Then
        MsgBox "输入格式错误，请检查！当前要求输入数字。", 48, "输入错误！"
        KeyAscii = 0
    End If
End If
End Sub
