VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "计算清单"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11730
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11730
   StartUpPosition =   2  '屏幕中心
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
      Height          =   585
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6840
      Width           =   1725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "查看装箱图"
      Height          =   540
      Left            =   9660
      TabIndex        =   22
      Top             =   6090
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确认并开始计算"
      Height          =   540
      Left            =   7665
      TabIndex        =   13
      Top             =   6090
      Width           =   1800
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   7455
      TabIndex        =   11
      Top             =   5040
      Width           =   4110
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   540
         Left            =   105
         TabIndex        =   12
         Top             =   240
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   953
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame3 
      Height          =   750
      Left            =   7455
      TabIndex        =   9
      Top             =   3990
      Width           =   4110
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "有X种装箱方案"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   315
         Width           =   3900
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "选择的装箱策略"
      Height          =   3585
      Left            =   7455
      TabIndex        =   5
      Top             =   105
      Width           =   4110
      Begin VB.OptionButton Option2 
         Caption         =   "深度搜索"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   21
         Top             =   3045
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.OptionButton Option2 
         Caption         =   "浅度搜索"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   1785
         TabIndex        =   20
         Top             =   3045
         Width           =   1485
      End
      Begin VB.CheckBox Check2 
         Caption         =   "深度搜索"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.CheckBox Check2 
         Caption         =   "浅度搜索"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   1785
         TabIndex        =   18
         Top             =   2100
         Value           =   1  'Checked
         Width           =   2010
      End
      Begin VB.CheckBox Check1 
         Caption         =   "宽大优先"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   17
         Top             =   840
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "长大优先"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   16
         Top             =   1260
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "数量优先"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   1785
         TabIndex        =   15
         Top             =   840
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "体积优先"
         Enabled         =   0   'False
         Height          =   225
         Index           =   3
         Left            =   1785
         TabIndex        =   14
         Top             =   1260
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "剩余空间拆分策略："
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   2625
         Width           =   3690
      End
      Begin VB.Label Label4 
         Caption         =   "工作面拆分策略："
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1680
         Width           =   3795
      End
      Begin VB.Label Label3 
         Caption         =   "装箱优先策略："
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   420
         Width           =   3690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "装箱列表"
      Height          =   7470
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   7155
      Begin MSComctlLib.ListView ListView2 
         Height          =   2400
         Left            =   210
         TabIndex        =   1
         Top             =   630
         Width           =   6690
         _ExtentX        =   11800
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
            Text            =   "自高(mm)"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   3840
         Left            =   210
         TabIndex        =   2
         Top             =   3465
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   6773
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
      Begin VB.Label Label2 
         Caption         =   "待装箱的货物："
         Height          =   330
         Left            =   210
         TabIndex        =   4
         Top             =   3255
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "选择的容器："
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   420
         Width           =   2535
      End
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   10920
      Picture         =   "Form4.frx":0ECA
      Top             =   6960
      Width           =   720
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'申皇专用
'****************************************************************************
Private containers(), bales() As String
Private m_count As Integer
Private dt_x As Boolean      '是否使用了缓冲器空间
Public Sub getcount()
k = CInt(ListView2.ListItems.Count)
i = 0
For Each check In Check1
    If check.Value = 1 Then i = i + 1
Next
j = 0
For Each check In Check2
    If check.Value = 1 Then j = j + 1
Next
m_count = k * i * j
Label6.Caption = Replace(Label6.Caption, "X", CStr(k * i * j))

End Sub

Private Sub Command1_Click()
'打开临时文件
Open App.Path + "\temps.txt" For Output As #1
Open App.Path + "\errs.txt" For Output As #2
'循环容器
pro_index = 1
For i = 1 To CInt(ListView2.ListItems.Count)
    
    container_index = i
    '循环优先策略
    For Each checks1 In Check1
        If checks1.Value = 1 Then
            order_flag = CInt(checks1.Index)
            pro_str2 = " |优先策略:" + checks1.Caption
            '循环工作面策略
            For Each checks2 In Check2
                If checks2.Value = 1 Then
                    If checks2.Index = 0 Then
                        WF_flag = 0 '深度搜索
                    Else
                        WF_flag = 1 '浅度搜索
                    End If
                    pro_str3 = " |工作面策略:" + checks2.Caption
                    '循环剩余空间策略
                        For Each checks3 In Option2
                            If checks3.Value = True Then
                                If checks3.Index = 0 Then
                                    SR_flag = 0 '深度搜索
                                Else
                                    SR_flag = 1 '浅度搜索
                                End If
                                pro_str4 = " |剩余空间策略:" + checks3.Caption
                                PRO_STR1 = "PRO_NUM|" + CStr(pro_index) + "|容器:" + ListView2.ListItems(i).SubItems(1)
                                Print #1, PRO_STR1 + pro_str2 + pro_str3 + pro_str4
                                '开始实际计算
                                Get_Container '向中间数组加载数据
                                Get_Bale
                                OrderBales order_flag '重排货物列表
                                '计算
                                ''Debug.Print "WF_flag=" + CStr(WF_flag)
                                ''Debug.Print "SR_flag=" + CStr(SR_flag)
                                ''Debug.Print "order_flag=" + CStr(order_flag)
                                ''Debug.Print "container_index=" + CStr(container_index)
                                Select Case WF_flag
                                Case 0
                                    
                                    a = Tests2(SR_flag, container_index)
                                Case 1
                                    a = Tests(SR_flag, container_index)
                                End Select
                                
                                ProgressBar1.Value = 100 * pro_index / m_count
                                
                                pro_index = pro_index + 1
                                DoEvents
                            End If
                        Next
                End If
            Next
        End If
    Next
Next i
'关闭文件
Close
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command12_Click()
Form3.Show

End Sub

'========================================================================
Private Sub Command2_Click()
Form5.Show 1
End Sub

Private Sub Form_Load()
Command2.Enabled = False
End Sub
'充填中间数组
Private Function Get_Container() As Boolean
ReDim containers(1 To CInt(ListView2.ListItems.Count), 0 To 7)
For i = 1 To CInt(ListView2.ListItems.Count)
    For j = 0 To 6
        containers(i, j) = ""
    Next j
Next i
i = 1
For Each items In ListView2.ListItems
    containers(i, 0) = "0"
    containers(i, 1) = items.SubItems(1) '名称
    containers(i, 2) = items.SubItems(2) '长
    containers(i, 3) = items.SubItems(3) '宽
    containers(i, 4) = items.SubItems(4) '高
    containers(i, 5) = CStr(CDbl(items.SubItems(2)) * CDbl(items.SubItems(3)) * CDbl(items.SubItems(4))) '体积
    containers(i, 7) = CStr(items.Index)
    i = i + 1
Next
End Function
'货物数据
Private Function Get_Bale() As Boolean
ReDim bales(1 To CInt(ListView3.ListItems.Count), 0 To 8)
For i = 1 To CInt(ListView3.ListItems.Count)
    For j = 0 To 7
        bales(i, j) = ""
    Next j
Next i
i = 1
For Each items In ListView3.ListItems
    bales(i, 0) = "0"
    bales(i, 1) = items.SubItems(1) '名称
    bales(i, 2) = items.SubItems(2) '长
    bales(i, 3) = items.SubItems(3) '宽
    bales(i, 4) = items.SubItems(4) '高
    bales(i, 5) = CStr(CDbl(items.SubItems(2)) * CDbl(items.SubItems(3)) * CDbl(items.SubItems(4))) '体积
    bales(i, 6) = ""
    bales(i, 7) = items.SubItems(6)
    bales(i, 8) = CStr(items.Index)
    i = i + 1
Next
End Function
'优先级排序 多种货物混装时候选择以什么标准为优先级
Private Sub OrderBales(ByVal types As Integer)
Dim tmps(1 To 8)
Select Case types
Case 0 '宽度排大先
    For i = 1 To UBound(bales, 1) - 1
        For j = i + 1 To UBound(bales, 1)
            If CDbl(bales(j, 3)) > CDbl(bales(i, 3)) Then
                For z = 1 To 8
                    tmps(z) = bales(i, z)
                    bales(i, z) = bales(j, z)
                    bales(j, z) = tmps(z)
                Next z
            End If
            
        Next j
    Next i
Case 1 '长度排长先
    For i = 1 To UBound(bales, 1) - 1
        For j = i + 1 To UBound(bales, 1)
            If CDbl(bales(j, 2)) > CDbl(bales(i, 2)) Then
                For z = 1 To 8
                    tmps(z) = bales(i, z)
                    bales(i, z) = bales(j, z)
                    bales(j, z) = tmps(z)
                Next z
            End If
            
        Next j
    Next i
Case 2 '体积排大先
    For i = 1 To UBound(bales, 1) - 1
        For j = i + 1 To UBound(bales, 1)
            If CDbl(bales(j, 5)) > CDbl(bales(i, 5)) Then
                For z = 1 To 8
                    tmps(z) = bales(i, z)
                    bales(i, z) = bales(j, z)
                    bales(j, z) = tmps(z)
                Next z
            End If
            
        Next j
    Next i
Case 3 '数量排多先
    For i = 1 To UBound(bales, 1) - 1
        For j = i + 1 To UBound(bales, 1)
            If CInt(bales(j, 7)) > CInt(bales(i, 7)) Then
                For z = 1 To 8
                    tmps(z) = bales(i, z)
                    bales(i, z) = bales(j, z)
                    bales(j, z) = tmps(z)
                Next z
            End If
        Next j
    Next i
End Select
End Sub

'小工作面
Private Function Tests(ByVal S_type As Integer, ByVal C_index As Integer) As Boolean
'S_type 剩余空间搜索策略, C_index 容器在容器数组中的索引
    '================定义变量===================
    Dim Max_X, Max_Y, Max_Z, WFace_X, WFace_Y, WFace_Z, CON_V, Bales_V, CON_EFF As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '货物的尺寸
    Dim CON_N, T_Flag_Ok, SNum, TempIndex, TempsIndexs, Bales_N, Bales_Whirl As Integer
    Dim Con_Index_Num, Bales_Index As String
    Dim i As Integer
    Dim Start_X, Start_Y, Start_Z, SMax_X As Double
    '工作面起点坐标
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3 As Double '当前工作面剩余空间的尺寸
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3 As Double '剩余空间起点坐标
    Dim W_Z_N, W_Y_N, W_Z_CON, Can_Count As Integer
    '===========================================
    '================初始化参数=================
    '读容器尺寸
    Max_X = CDbl(containers(C_index, 2))
    Max_Y = CDbl(containers(C_index, 3))
    Max_Z = CDbl(containers(C_index, 4))
    Con_Index_Num = containers(C_index, 7) '容器在LISTVIEW2中的LISTITEM编号
    CON_V = 0 '容器总体积
    CON_N = 1 '使用的容器数量
    T_Flag_Ok = 0 '装载完毕标志 0 未完成 1完成
    Bales_V = 0 '货物总体积
    '读取货物总体积
    For i = 1 To UBound(bales, 1)
        Bales_V = Bales_V + CDbl(bales(i, 5)) * CInt(bales(i, 7))
    Next i
    '===========================================
    '================开始装箱计算===============
    Do '容器循环
        SNum = 0 '工作面编号
        '当前工作面起点坐标
        Start_X = 0
        Start_Y = 0
        Start_Z = 0
        dt_x = False
        '设置工作空间最大尺寸等于容器尺寸
        WFace_X = Max_X
        WFace_Y = Max_Y
        WFace_Z = Max_Z
        SMax_X = Max_X ' X方向上可用的最大尺寸
        '工作面循环开始
        Do
            '设置参数
            TempIndex = 0 '选择的货物在货物数组中的索引
            '设置当前工作面可用空间尺寸
            WFace_Y = Max_Y
            WFace_Z = Max_Z
            WFace_X = SMax_X
            '选择箱子先进行装载
            For i = 1 To UBound(bales, 1) '遍历货物数组搜寻可以放入工作面空间的货物
                If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '是否可装载，并且剩余数量大于0
                    '选中进行装箱的货物的尺寸
                    Bales_X = CDbl(bales(i, 2))
                    Bales_Y = CDbl(bales(i, 3))
                    Bales_Z = CDbl(bales(i, 4))
                    '选中进行装箱的货物数量
                    Bales_N = CInt(bales(i, 7))
                    Bales_Index = bales(i, 8) '选中的货物在LISTVIEW3中的LISTITEM编号
                    Bales_Whirl = 0 '设置货物是否水平旋转标志 0 未旋转 1 旋转90度 Bales_X、Bales_Y互换
                    '缓冲器调节长度
                    '如果货物长度大于工作面可用长度但小于工作面可用长度加缓冲器的长度
                    If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False And hc Then  'New Code
                        dt_x = True '缓冲器已使用
                        WFace_X = WFace_X + 100 '设置工作面可用长度等于工作面可用长度加缓冲器长度
                        TempIndex = i '设置货物在货物数组中的索引编号
                        Exit For '退出选择货物循环
                    ElseIf Bales_X < WFace_X Then '货物长度小于工作面可用长度
                        TempIndex = i
                        Exit For
                    End If
                End If
            Next i
            '如果没有合适装载尺寸的货物，货物水平旋转
            If TempIndex = 0 Then
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then  '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
                        Bales_Y = CDbl(bales(i, 2))
                        Bales_X = CDbl(bales(i, 3))
                        Bales_Z = CDbl(bales(i, 4))
                        Bales_N = CInt(bales(i, 7))
                        Bales_Index = bales(i, 8)
                        Bales_Whirl = 1
                        '缓冲器调节长度
                        If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False And hc Then  'New Code
                            dt_x = True
                            WFace_X = WFace_X + 100
                            TempIndex = i
                            'Text2.Text = Text2.Text + CStr(tempindex) + vbCrLf
                            Exit For
                        ElseIf Bales_X < WFace_X Then
                            TempIndex = i
                            'Text2.Text = Text2.Text + CStr(tempindex) + vbCrLf
                            Exit For
                        End If
                    End If
                Next i
            End If
            '如果还没有可装载的货物
            If TempIndex = 0 Then
                '判断是否装箱完成
                TempsIndexs = 0
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 Then '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
                        TempsIndexs = i
                        Exit For
                    End If
                Next i
                If TempsIndexs = 0 Then
                    Tests = True '装箱成功
                    T_Flag_Ok = 1
                    Exit Do '装箱成功
                Else
                    '当前容器剩余空间无法放下合适的箱子,使用下一个容器
                    Exit Do '退出工作面循环
                End If
            End If
            SNum = SNum + 1 '设置工作面编号
            

     '宽度方向上可装入的数量
     W_Y_N = Int(WFace_Y / Bales_Y)
    '垂直方向上可装入的数量
    W_Z_N = Int(WFace_Z / Bales_Z)
    '工作面可装入箱子的数量是否小于箱子总数
    If W_Y_N * W_Z_N > Bales_N Then
        Can_Count = Bales_N '已装载的箱子数
    Else
        Can_Count = W_Y_N * W_Z_N
    End If
    If (Can_Count Mod W_Y_N) > 0 Then '在Z方向需装载的行数
        If Can_Count < W_Y_N Then
            W_Z_CON = 1
        Else
            W_Z_CON = Int(Can_Count / W_Y_N) + 1
        End If
    Else
        W_Z_CON = Int(Can_Count / W_Y_N)
    End If

        S_Start_X1 = Start_X
        S_Start_X2 = Start_X
        S_Start_X3 = Start_X
    If (Can_Count Mod W_Y_N) = 0 Then '无剩余空间2
        Re_X1 = Bales_X
        Re_X2 = 0
        Re_X3 = Bales_X
        Re_Y1 = W_Y_N * Bales_Y
        Re_Y2 = 0
        Re_Y3 = WFace_Y - Bales_Y * W_Y_N
        Re_Z1 = WFace_Z - Bales_Z * W_Z_CON
        Re_Z2 = 0
        Re_Z3 = WFace_Z
        S_Start_Y1 = Start_Y
        S_Start_Y2 = Start_Y + W_Y_N * Bales_Y
        S_Start_Y3 = Start_Y + W_Y_N * Bales_Y
        S_Start_Z1 = Start_Z + Bales_Z * W_Z_CON
        S_Start_Z2 = Start_Z
        S_Start_Z3 = Start_Z
    Else
        If W_Z_CON > 1 Then
            Re_X1 = Bales_X
            Re_X2 = Bales_X
            Re_X3 = Bales_X
            Re_Y1 = (Can_Count Mod W_Y_N) * Bales_Y
            Re_Y2 = (W_Y_N - (Can_Count Mod W_Y_N)) * Bales_Y
            Re_Y3 = WFace_Y - Bales_Y * W_Y_N
            Re_Z1 = WFace_Z - Bales_Z * W_Z_CON
            Re_Z2 = WFace_Z - Bales_Z * (W_Z_CON - 1)
            Re_Z3 = WFace_Z
            S_Start_Y1 = Start_Y
            S_Start_Y2 = Start_Y + (Can_Count Mod W_Y_N) * Bales_Y
            S_Start_Y3 = Start_Y + Bales_Y * W_Y_N
            S_Start_Z1 = Start_Z + Bales_Z * W_Z_CON
            S_Start_Z2 = Start_Z + Bales_Z * (W_Z_CON - 1)
            S_Start_Z3 = Start_Z
        Else '无剩余空间3
            Re_X1 = Bales_X
            Re_X2 = Bales_X
            RZ_X3 = 0
            Re_Y1 = (Can_Count Mod W_Y_N) * Bales_Y
            Re_Y2 = WFace_Y - (Can_Count Mod W_Y_N) * Bales_Y
            Re_Y3 = 0
            Re_Z1 = WFace_Z - Bales_Z * W_Z_CON
            Re_Z2 = WFace_Z
            Re_Z3 = 0
            S_Start_Y1 = Start_Y
            S_Start_Y2 = Start_Y + (Can_Count Mod W_Y_N) * Bales_Y
            S_Start_Y3 = Start_Y + (Can_Count Mod W_Y_N) * Bales_Y
            S_Start_Z1 = Start_Z + Bales_Z * W_Z_CON
            S_Start_Z2 = Start_Z
            S_Start_Z3 = Start_Z
        End If
    End If

    'debugstr = "CON:" + CStr(CON_N) + " WF:" + CStr(SNum) + vbCrLf
    'debugstr = debugstr + "RX1:" + CStr(Re_X1) + " RX2:" + CStr(Re_X2) + " RX3:" + CStr(Re_X3) + vbCrLf
    'debugstr = debugstr + "RY1:" + CStr(Re_Y1) + " RY2:" + CStr(Re_Y2) + " RY3:" + CStr(Re_Y3) + vbCrLf
    'debugstr = debugstr + "RZ1:" + CStr(Re_Z1) + " RZ2:" + CStr(Re_Z2) + " RZ3:" + CStr(Re_Z3) + vbCrLf
    'debugstr = debugstr + "SX1:" + CStr(S_Start_X1) + " SX2:" + CStr(S_Start_X2) + " SX3:" + CStr(S_Start_X3) + vbCrLf
    'debugstr = debugstr + "SY1:" + CStr(S_Start_Y1) + " SY2:" + CStr(S_Start_Y2) + " SY3:" + CStr(S_Start_Y3) + vbCrLf
    'debugstr = debugstr + "SZ1:" + CStr(S_Start_Z1) + " SZ2:" + CStr(S_Start_Z2) + " SZ3:" + CStr(S_Start_Z3) + vbCrLf
    'debugstr = debugstr + "WY:" + CStr(W_Y_N) + " WZ:" + CStr(W_Z_N) + " WC:" + CStr(Can_Count) + vbCrLf
    'Print #2, debugstr
            bales(TempIndex, 0) = "1"
            bales(TempIndex, 7) = CStr(CInt(bales(TempIndex, 7)) - Can_Count)
            SMax_X = SMax_X - Bales_X '剩余长度方向可用尺寸
            '计算装箱图
            '第N个容器|容器在LISTVIEW中索引|工作面编号|子工作面编号|货物再LISTVIEW中索引
            '|货物名称|装载的货物数量|X方向可装载的数量|Y方向可装载的数量|Z方向可装载的数量
            '|起点X坐标|起点Y坐标|起点Z坐标|是否水平旋转
            tempsstr = CStr(CON_N) + "|" + Con_Index_Num + "|" + CStr(SNum) + "|" + CStr("0") + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
            tempsstr = tempsstr + CStr(1) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(Start_X) + "|" + CStr(Start_Y) + "|" + CStr(Start_Z) + "|" + CStr(Bales_Whirl)
            Print #1, tempsstr            'Debug.Print "中间参数 " + CStr(snum) + " " + CStr(S_type)
            'Debug.Print "浅度搜索 打印工作面装箱列表"
            '剩余空间充填
            If S_type = 0 Then '采用深度搜索
                'Debug.Print "浅度搜索 剩余深度搜索"
                Respace2 Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace2 Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace2 Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
            Else '采用浅度搜索
                'Debug.Print "浅度搜索 剩余浅度搜索"
                Respace Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
            End If

            '下一工作面起点坐标
            Start_X = Start_X + Bales_X
            Start_Y = 0
            Start_Z = 0
            DoEvents
        Loop
        If T_Flag_Ok = 1 Then
            CON_V = CON_V + (containers(C_index, 2) - SMax_X) * containers(C_index, 3) * containers(C_index, 4)
            Exit Do
        Else
            CON_N = CON_N + 1
            If dt_x = False Then
                CON_V = CON_V + containers(C_index, 2) * containers(C_index, 3) * containers(C_index, 4)
            Else
                CON_V = CON_V + (containers(C_index, 2) + 100) * containers(C_index, 3) * containers(C_index, 4) '利用了缓冲器后的容器体积
            End If
        End If
    Loop
    '计算效率
    CON_EFF = Bales_V / CON_V
    Print #1, "PRO_EFF=|" + CStr(CON_EFF) + "| CON_N=" + CStr(CON_N)
End Function
'大工作面
Private Function Tests2(ByVal S_type As Integer, ByVal C_index As Integer) As Boolean
'S_type 剩余空间搜索策略, C_index 容器在容器数组中的索引
    '================定义变量===================
    Dim Max_X, Max_Y, Max_Z, WFace_X, WFace_Y, WFace_Z, CON_V, Bales_V, CON_EFF As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '货物的尺寸
    Dim CON_N, T_Flag_Ok, SNum, TempIndex, TempsIndexs, Bales_N, Bales_Whirl As Integer
    Dim Con_Index_Num, Bales_Index As String
    Dim i As Integer
    Dim Start_X, Start_Y, Start_Z, SMax_X As Double
    '工作面起点坐标
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3, Re_X4, Re_Y4, Re_Z4 As Double '当前工作面剩余空间的尺寸
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3, S_Start_X4, S_Start_Y4, S_Start_Z4 As Double '剩余空间起点坐标
    Dim W_Z_N, W_Y_N, W_X_N, W_X_CON, Can_Count As Integer
    '===========================================
    '================初始化参数=================
    '读容器尺寸
    Max_X = CDbl(containers(C_index, 2))
    Max_Y = CDbl(containers(C_index, 3))
    Max_Z = CDbl(containers(C_index, 4))
    Con_Index_Num = containers(C_index, 7) '容器在LISTVIEW2中的LISTITEM编号
    CON_V = 0 '容器总体积
    CON_N = 1 '使用的容器数量
    T_Flag_Ok = 0 '装载完毕标志 0 未完成 1完成
    Bales_V = 0 '货物总体积
    '读取货物总体积
    For i = 1 To UBound(bales, 1)
        Bales_V = Bales_V + CDbl(bales(i, 5)) * CInt(bales(i, 7))
    Next i
    '===========================================
    '================开始装箱计算===============
    Do '容器循环
        SNum = 0 '工作面编号
        '当前工作面起点坐标
        Start_X = 0
        Start_Y = 0
        Start_Z = 0
        dt_x = False
        '设置工作空间最大尺寸等于容器尺寸
        WFace_X = Max_X
        WFace_Y = Max_Y
        WFace_Z = Max_Z
        SMax_X = Max_X ' X方向上可用的最大尺寸
        '工作面循环开始
        Do
            '设置参数
            TempIndex = 0 '选择的货物在货物数组中的索引
            '设置当前工作面可用空间尺寸
            WFace_Y = Max_Y
            WFace_Z = Max_Z
            WFace_X = SMax_X
            '选择箱子先进行装载
            For i = 1 To UBound(bales, 1) '遍历货物数组搜寻可以放入工作面空间的货物
                If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '是否可装载，并且剩余数量大于0
                    '选中进行装箱的货物的尺寸
                    Bales_X = CDbl(bales(i, 2))
                    Bales_Y = CDbl(bales(i, 3))
                    Bales_Z = CDbl(bales(i, 4))
                    '选中进行装箱的货物数量
                    Bales_N = CInt(bales(i, 7))
                    Bales_Index = bales(i, 8) '选中的货物在LISTVIEW3中的LISTITEM编号
                    Bales_Whirl = 0 '设置货物是否水平旋转标志 0 未旋转 1 旋转90度 Bales_X、Bales_Y互换
                    '缓冲器调节长度
                    '如果货物长度大于工作面可用长度但小于工作面可用长度加缓冲器的长度
                    If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False And hc Then 'New Code
                        dt_x = True '缓冲器已使用
                        WFace_X = WFace_X + 100 '设置工作面可用长度等于工作面可用长度加缓冲器长度
                        TempIndex = i '设置货物在货物数组中的索引编号
                        Exit For '退出选择货物循环
                    ElseIf Bales_X < WFace_X Then '货物长度小于工作面可用长度
                        TempIndex = i
                        Exit For
                    End If
                End If
            Next i
            '如果没有合适装载尺寸的货物，货物水平旋转
            If TempIndex = 0 Then
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then  '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
                        Bales_Y = CDbl(bales(i, 2))
                        Bales_X = CDbl(bales(i, 3))
                        Bales_Z = CDbl(bales(i, 4))
                        Bales_N = CInt(bales(i, 7))
                        Bales_Index = bales(i, 8)
                        Bales_Whirl = 1
                        '缓冲器调节长度
                        If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False And hc Then  'New Code
                            dt_x = True
                            WFace_X = WFace_X + 100
                            TempIndex = i

                            Exit For
                        ElseIf Bales_X < WFace_X Then
                            TempIndex = i

                            Exit For
                        End If
                    End If
                Next i
            End If
            '如果还没有可装载的货物
            If TempIndex = 0 Then
                '判断是否装箱完成
                TempsIndexs = 0
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 Then '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
                        TempsIndexs = i
                        Exit For
                    End If
                Next i
                If TempsIndexs = 0 Then
                    Tests2 = True '装箱成功
                    T_Flag_Ok = 1
                    Exit Do '装箱成功
                Else
                    '当前容器剩余空间无法放下合适的箱子,使用下一个容器
                    Exit Do '退出工作面循环
                End If
            End If
            SNum = SNum + 1 '设置工作面编号
            '宽度方向上可装入的数量
            W_Y_N = Int(WFace_Y / Bales_Y)
            '垂直方向上可装入的数量
            W_Z_N = Int(WFace_Z / Bales_Z)
            '长度方向上可装入的数量
            W_X_N = Int(WFace_X / Bales_X)
            '工作面可装入箱子的数量是否小于箱子总数
            If W_Y_N * W_Z_N * W_X_N > Bales_N Then
                Can_Count = Bales_N
            Else
                Can_Count = W_Y_N * W_Z_N * W_X_N
            End If
            '情况1
            If (Can_Count Mod (W_Y_N * W_Z_N)) = 0 Then
                W_X_CON = Can_Count / (W_Y_N * W_Z_N)
                Re_X1 = Bales_X * W_X_CON
                Re_X2 = Bales_X * W_X_CON
                Re_X3 = 0
                Re_X4 = 0
                Re_Y1 = Bales_Y * W_Y_N
                Re_Y2 = WFace_Y - Bales_Y * W_Y_N
                Re_Y3 = 0
                Re_Y4 = 0
                Re_Z1 = WFace_Z - Bales_Z * W_Z_N
                Re_Z2 = WFace_Z
                Re_Z3 = 0
                Re_Z4 = 0
                S_Start_X1 = Start_X
                S_Start_X2 = Start_X
                S_Start_X3 = 0
                S_Start_X4 = 0
                S_Start_Y1 = Start_Y
                S_Start_Y2 = Start_Y + Bales_Y * W_Y_N
                S_Start_Y3 = 0
                S_Start_Y4 = 0
                S_Start_Z1 = Start_Z + Bales_Z * W_Z_N
                S_Start_Z2 = Start_Z
                S_Start_Z3 = 0
                S_Start_Z4 = 0
            Else
                W_X_CON = Int(Can_Count / (W_Y_N * W_Z_N)) + 1
                Re_X1 = Bales_X * (W_X_CON - 1)
                Re_X2 = Bales_X * W_X_CON
                Re_Y1 = Bales_Y * W_Y_N
                Re_Y2 = WFace_Y - Bales_Y * W_Y_N
                Re_Z1 = WFace_Z - Bales_Z * W_Z_N
                Re_Z2 = WFace_Z
                '剩余空间起点坐标
                S_Start_X1 = Start_X
                S_Start_X2 = Start_X
                S_Start_Y1 = Start_Y
                S_Start_Y2 = Start_Y + Bales_Y * W_Y_N
                S_Start_Z1 = Start_Z + Bales_Z * W_Z_N
                S_Start_Z2 = Start_Z
                If ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N) > 0 Then
                    Re_X3 = Bales_X
                    Re_X4 = Bales_X
                    Re_Y3 = Bales_Y * ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N)
                    Re_Y4 = Bales_Y * (W_Y_N - ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N))
                    Re_Z3 = WFace_Z - Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N) + 1)
                    Re_Z4 = WFace_Z - Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                    S_Start_X3 = Start_X + Bales_X * (W_X_CON - 1)
                    S_Start_X4 = Start_X + Bales_X * (W_X_CON - 1)
                    S_Start_Y3 = Start_Y
                    S_Start_Y4 = Start_Y + Bales_Y * ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N)
                    S_Start_Z3 = Start_Z + Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N) + 1)
                    S_Start_Z4 = Start_Z + Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                Else
                    Re_X3 = Bales_X
                    Re_X4 = 0
                    Re_Y3 = Bales_Y * W_Y_N
                    Re_Y4 = 0
                    Re_Z3 = WFace_Z - Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                    Re_Z4 = 0
                    S_Start_X3 = Start_X + Bales_X * (W_X_CON - 1)
                    S_Start_X4 = 0
                    S_Start_Y3 = Start_Y
                    S_Start_Y4 = 0
                    S_Start_Z3 = Start_Z + Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                    S_Start_Z4 = 0
                End If
            End If
            bales(TempIndex, 0) = "1"
            bales(TempIndex, 7) = CStr(CInt(bales(TempIndex, 7)) - Can_Count)
            SMax_X = SMax_X - Bales_X * W_X_CON '剩余长度方向可用尺寸
            tempsstr = CStr(CON_N) + "|" + Con_Index_Num + "|" + CStr(SNum) + "|" + CStr("0") + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
            tempsstr = tempsstr + CStr(W_X_CON) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(Start_X) + "|" + CStr(Start_Y) + "|" + CStr(Start_Z) + "|" + CStr(Bales_Whirl)
            Print #1, tempsstr
            '剩余空间充填
            If S_type = 0 Then '采用深度搜索
                Respace2 Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace2 Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace2 Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
                Respace2 Con_Index_Num, CON_N, S_type, Re_X4, Re_Y4, Re_Z4, SNum, 4, S_Start_X4, S_Start_Y4, S_Start_Z4
            Else '采用浅度搜索
                Respace Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
                Respace Con_Index_Num, CON_N, S_type, Re_X4, Re_Y4, Re_Z4, SNum, 4, S_Start_X4, S_Start_Y4, S_Start_Z4
            End If
            '下一工作面起点坐标
            Start_X = Start_X + Bales_X * W_X_CON
            Start_Y = 0
            Start_Z = 0
            DoEvents
        Loop
        'Print #2, "TEMP 1"
        If T_Flag_Ok = 1 Then
            CON_V = CON_V + (containers(C_index, 2) - SMax_X) * containers(C_index, 3) * containers(C_index, 4)
            Exit Do
        Else
            CON_N = CON_N + 1
            If dt_x = False Then
                CON_V = CON_V + containers(C_index, 2) * containers(C_index, 3) * containers(C_index, 4)
            Else
                CON_V = CON_V + (containers(C_index, 2) + 100) * containers(C_index, 3) * containers(C_index, 4) '利用了缓冲器后的容器体积
            End If
        End If
        DoEvents
    Loop
    '计算效率
    CON_EFF = Bales_V / CON_V
    Print #1, "PRO_EFF=|" + CStr(CON_EFF) + "| CON_N=" + CStr(CON_N)
End Function
'充填剩余空间
Private Sub Respace(ByVal CON_I_N As String, ByVal CON_NS As Integer, ByVal S_type As Integer, x_max, y_max, z_max, snums, Index, SStart_X, SStart_Y, SStart_Z)
    '================定义变量===================
    Dim WFace_X, WFace_Y, WFace_Z As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '货物的尺寸
    Dim TempIndex, Bales_N, Bales_Whirl As Integer
    Dim Bales_Index As String
    Dim i As Integer
    '工作面起点坐标
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3 As Double '当前工作面剩余空间的尺寸
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3 As Double '剩余空间起点坐标
    Dim W_Z_N, W_Y_N, W_Z_CON, Can_Count As Integer
    '===========================================
    TempIndex = 0
    WFace_Y = y_max
    WFace_Z = z_max
    WFace_X = x_max
    '选择箱子先进行装载
    
    For i = 1 To UBound(bales, 1)
        If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
            Bales_X = CDbl(bales(i, 2))
            Bales_Y = CDbl(bales(i, 3))
            Bales_Z = CDbl(bales(i, 4))
            Bales_N = CInt(bales(i, 7))
            Bales_Index = bales(i, 8)
            Bales_Whirl = 0
            '缓冲器调节长度
            'If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False Then
            '    dt_x = True
            '    WFace_X = WFace_X + 100
            '    TempIndex = I
            '    Exit For
            'Else
            If Bales_X < WFace_X Then
                TempIndex = i
                Exit For
            End If
        End If
    Next i
    '如果没有合适装载尺寸的货物，货物水平旋转
    If TempIndex = 0 Then
        For i = 1 To UBound(bales, 1)
            If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then   '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
                Bales_Y = CDbl(bales(i, 2))
                Bales_X = CDbl(bales(i, 3))
                Bales_Z = CDbl(bales(i, 4))
                Bales_N = CInt(bales(i, 7))
                Bales_Index = bales(i, 8)
                Bales_Whirl = 1
                '缓冲器调节长度
                'If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False Then
                '    dt_x = True
                '    WFace_X = WFace_X + 100
                '    TempIndex = I
                '    Exit For
                'Else
                If Bales_X < WFace_X Then
                    TempIndex = i
                    Exit For
                End If
            End If
        Next i
    End If
    If TempIndex = 0 Then Exit Sub
    '宽度方向上可装入的数量
     W_Y_N = Int(WFace_Y / Bales_Y)
    '垂直方向上可装入的数量
    W_Z_N = Int(WFace_Z / Bales_Z)
    '工作面可装入箱子的数量是否小于箱子总数
    If W_Y_N * W_Z_N > Bales_N Then
        Can_Count = Bales_N '已装载的箱子数
    Else
        Can_Count = W_Y_N * W_Z_N
    End If
    If (Can_Count Mod W_Y_N) > 0 Then '在Z方向需装载的行数
        If Can_Count < W_Y_N Then
            W_Z_CON = 1
        Else
            W_Z_CON = Int(Can_Count / W_Y_N) + 1
        End If
    Else
        W_Z_CON = Int(Can_Count / W_Y_N)
    End If

        S_Start_X1 = SStart_X
        S_Start_X2 = SStart_X
        S_Start_X3 = SStart_X
    If (Can_Count Mod W_Y_N) = 0 Then '无剩余空间2
        Re_X1 = Bales_X
        Re_X2 = 0
        Re_X3 = Bales_X
        Re_Y1 = W_Y_N * Bales_Y
        Re_Y2 = 0
        Re_Y3 = WFace_Y - Bales_Y * W_Y_N
        Re_Z1 = WFace_Z - Bales_Z * W_Z_CON
        Re_Z2 = 0
        Re_Z3 = WFace_Z
        S_Start_Y1 = SStart_Y
        S_Start_Y2 = SStart_Y + W_Y_N * Bales_Y
        S_Start_Y3 = SStart_Y + W_Y_N * Bales_Y
        S_Start_Z1 = SStart_Z + Bales_Z * W_Z_CON
        S_Start_Z2 = SStart_Z
        S_Start_Z3 = SStart_Z
    Else
        If W_Z_CON > 1 Then
            Re_X1 = Bales_X
            Re_X2 = Bales_X
            Re_X3 = Bales_X
            Re_Y1 = (Can_Count Mod W_Y_N) * Bales_Y
            Re_Y2 = (W_Y_N - (Can_Count Mod W_Y_N)) * Bales_Y
            Re_Y3 = WFace_Y - Bales_Y * W_Y_N
            Re_Z1 = WFace_Z - Bales_Z * W_Z_CON
            Re_Z2 = WFace_Z - Bales_Z * (W_Z_CON - 1)
            Re_Z3 = WFace_Z
            S_Start_Y1 = SStart_Y
            S_Start_Y2 = SStart_Y + (Can_Count Mod W_Y_N) * Bales_Y
            S_Start_Y3 = SStart_Y + Bales_Y * W_Y_N
            S_Start_Z1 = SStart_Z + Bales_Z * W_Z_CON
            S_Start_Z2 = SStart_Z + Bales_Z * (W_Z_CON - 1)
            S_Start_Z3 = SStart_Z
        Else '无剩余空间3
            Re_X1 = Bales_X
            Re_X2 = Bales_X
            RZ_X3 = 0
            Re_Y1 = (Can_Count Mod W_Y_N) * Bales_Y
            Re_Y2 = WFace_Y - (Can_Count Mod W_Y_N) * Bales_Y
            Re_Y3 = 0
            Re_Z1 = WFace_Z - Bales_Z * W_Z_CON
            Re_Z2 = WFace_Z
            Re_Z3 = 0
            S_Start_Y1 = SStart_Y
            S_Start_Y2 = SStart_Y + (Can_Count Mod W_Y_N) * Bales_Y
            S_Start_Y3 = SStart_Y + (Can_Count Mod W_Y_N) * Bales_Y
            S_Start_Z1 = SStart_Z + Bales_Z * W_Z_CON
            S_Start_Z2 = SStart_Z
            S_Start_Z3 = SStart_Z
        End If
    End If

    'debugstr = "CON:" + CStr(CON_NS) + " SWF:" + CStr(Snums) + vbCrLf
    'debugstr = debugstr + "RX1:" + CStr(Re_X1) + " RX2:" + CStr(Re_X2) + " RX3:" + CStr(Re_X3) + vbCrLf
    'debugstr = debugstr + "RY1:" + CStr(Re_Y1) + " RY2:" + CStr(Re_Y2) + " RY3:" + CStr(Re_Y3) + vbCrLf
    'debugstr = debugstr + "RZ1:" + CStr(Re_Z1) + " RZ2:" + CStr(Re_Z2) + " RZ3:" + CStr(Re_Z3) + vbCrLf
    'debugstr = debugstr + "SX1:" + CStr(S_Start_X1) + " SX2:" + CStr(S_Start_X2) + " SX3:" + CStr(S_Start_X3) + vbCrLf
    'debugstr = debugstr + "SY1:" + CStr(S_Start_Y1) + " SY2:" + CStr(S_Start_Y2) + " SY3:" + CStr(S_Start_Y3) + vbCrLf
    'debugstr = debugstr + "SZ1:" + CStr(S_Start_Z1) + " SZ2:" + CStr(S_Start_Z2) + " SZ3:" + CStr(S_Start_Z3) + vbCrLf
    'debugstr = debugstr + "WY:" + CStr(W_Y_N) + " WZ:" + CStr(W_Z_N) + " WC:" + CStr(Can_Count) + vbCrLf
    'Print #2, debugstr
    bales(TempIndex, 0) = "1"
    bales(TempIndex, 7) = CStr(CInt(bales(TempIndex, 7)) - Can_Count)
    '计算装箱图
    tempsstr = CStr(CON_NS) + "|" + CON_I_N + "|" + CStr(snums) + "|" + CStr(Index) + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
    tempsstr = tempsstr + CStr(1) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(SStart_X) + "|" + CStr(SStart_Y) + "|" + CStr(SStart_Z) + "|" + CStr(Bales_Whirl)
    Print #1, tempsstr
    '剩余空间充填
    If S_type = 0 Then '采用深度搜索
        Respace2 CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace2 CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace2 CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
    Else '采用浅度搜索
        Respace CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
    End If
End Sub
Private Sub Respace2(ByVal CON_I_N As String, ByVal CON_NS As Integer, ByVal S_type As Integer, x_max, y_max, z_max, snums, Index, SStart_X, SStart_Y, SStart_Z)
    '================定义变量===================
    Dim WFace_X, WFace_Y, WFace_Z As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '货物的尺寸
    Dim TempIndex, Bales_N, Bales_Whirl As Integer
    Dim Bales_Index As String
    Dim i As Integer
    '工作面起点坐标
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3, Re_X4, Re_Y4, Re_Z4 As Double '当前工作面剩余空间的尺寸
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3, S_Start_X4, S_Start_Y4, S_Start_Z4 As Double '剩余空间起点坐标
    Dim W_Z_N, W_Y_N, W_X_N, W_X_CON, Can_Count As Integer
    '===========================================
    TempIndex = 0
    WFace_Y = y_max
    WFace_Z = z_max
    WFace_X = x_max
    '选择箱子先进行装载
    For i = 1 To UBound(bales, 1)
        If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
            Bales_X = CDbl(bales(i, 2))
            Bales_Y = CDbl(bales(i, 3))
            Bales_Z = CDbl(bales(i, 4))
            Bales_N = CInt(bales(i, 7))
            Bales_Index = bales(i, 8)
            Bales_Whirl = 0
            '缓冲器调节长度
            'If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False Then
            '    dt_x = True
            '    WFace_X = WFace_X + 100
            '    TempIndex = I
            '    Exit For
            'Else
            If Bales_X < WFace_X Then
                TempIndex = i
                Exit For
            End If
        End If
    Next i
    '如果没有合适装载尺寸的货物，货物水平旋转
    If TempIndex = 0 Then
        For i = 1 To UBound(bales, 1)
            If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then   '是否未装载，并且剩余数量大于0 Bales(i, 0) = 0 And
                Bales_Y = CDbl(bales(i, 2))
                Bales_X = CDbl(bales(i, 3))
                Bales_Z = CDbl(bales(i, 4))
                Bales_N = CInt(bales(i, 7))
                Bales_Index = bales(i, 8)
                Bales_Whirl = 1
                '缓冲器调节长度
                'If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False Then
                '    dt_x = True
                '    WFace_X = WFace_X + 100
                '    TempIndex = I
                '    Exit For
                'Else
                If Bales_X < WFace_X Then
                    TempIndex = i
                    Exit For
                End If
            End If
        Next i
    End If
    If TempIndex = 0 Then Exit Sub
    '宽度方向上可装入的数量
    W_Y_N = Int(WFace_Y / Bales_Y)
    '垂直方向上可装入的数量
    W_Z_N = Int(WFace_Z / Bales_Z)
    '长度方向上可装入的数量
    W_X_N = Int(WFace_X / Bales_X)
    '工作面可装入箱子的数量是否小于箱子总数
    If W_Y_N * W_Z_N * W_X_N > Bales_N Then
        Can_Count = Bales_N
    Else
        Can_Count = W_Y_N * W_Z_N * W_X_N
    End If
    If (Can_Count Mod (W_Y_N * W_Z_N)) = 0 Then
        W_X_CON = Can_Count / (W_Y_N * W_Z_N)
        Re_X1 = Bales_X * W_X_CON
        Re_X2 = Bales_X * W_X_CON
        Re_X3 = 0
        Re_X4 = 0
        Re_Y1 = Bales_Y * W_Y_N
        Re_Y2 = WFace_Y - Bales_Y * W_Y_N
        Re_Y3 = 0
        Re_Y4 = 0
        Re_Z1 = WFace_Z - Bales_Z * W_Z_N
        Re_Z2 = WFace_Z
        Re_Z3 = 0
        Re_Z4 = 0
        S_Start_X1 = SStart_X
        S_Start_X2 = SStart_X
        S_Start_X3 = 0
        S_Start_X4 = 0
        S_Start_Y1 = SStart_Y
        S_Start_Y2 = SStart_Y + Bales_Y * W_Y_N
        S_Start_Y3 = 0
        S_Start_Y4 = 0
        S_Start_Z1 = SStart_Z + Bales_Z * W_Z_N
        S_Start_Z2 = SStart_Z
        S_Start_Z3 = 0
        S_Start_Z4 = 0
    Else
        W_X_CON = Int(Can_Count / (W_Y_N * W_Z_N)) + 1
        Re_X1 = Bales_X * (W_X_CON - 1)
        Re_X2 = Bales_X * W_X_CON
        Re_Y1 = Bales_Y * W_Y_N
        Re_Y2 = WFace_Y - Bales_Y * W_Y_N
        Re_Z1 = WFace_Z - Bales_Z * W_Z_N
        Re_Z2 = WFace_Z
        '剩余空间起点坐标
        S_Start_X1 = SStart_X
        S_Start_X2 = SStart_X
        S_Start_Y1 = SStart_Y
        S_Start_Y2 = SStart_Y + Bales_Y * W_Y_N
        S_Start_Z1 = SStart_Z + Bales_Z * W_Z_N
        S_Start_Z2 = SStart_Z
                If ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N) > 0 Then
                    Re_X3 = Bales_X
                    Re_X4 = Bales_X
                    Re_Y3 = Bales_Y * ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N)
                    Re_Y4 = Bales_Y * (W_Y_N - ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N))
                    Re_Z3 = WFace_Z - Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N) + 1)
                    Re_Z4 = WFace_Z - Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                    S_Start_X3 = SStart_X + Bales_X * (W_X_CON - 1)
                    S_Start_X4 = SStart_X + Bales_X * (W_X_CON - 1)
                    S_Start_Y3 = SStart_Y
                    S_Start_Y4 = SStart_Y + Bales_Y * ((Can_Count Mod (W_Y_N * W_Z_N)) Mod W_Y_N)
                    S_Start_Z3 = SStart_Z + Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N) + 1)
                    S_Start_Z4 = SStart_Z + Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                Else
                    Re_X3 = Bales_X
                    Re_X4 = 0
                    Re_Y3 = Bales_Y * W_Y_N
                    Re_Y4 = 0
                    Re_Z3 = WFace_Z - Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                    Re_Z4 = 0
                    S_Start_X3 = SStart_X + Bales_X * (W_X_CON - 1)
                    S_Start_X4 = 0
                    S_Start_Y3 = SStart_Y
                    S_Start_Y4 = 0
                    S_Start_Z3 = SStart_Z + Bales_Z * (Int((Can_Count Mod (W_Y_N * W_Z_N)) / W_Y_N))
                    S_Start_Z4 = 0
                End If
    End If
    bales(TempIndex, 0) = "1"
    bales(TempIndex, 7) = CStr(CInt(bales(TempIndex, 7)) - Can_Count)
    tempsstr = CStr(CON_NS) + "|" + CON_I_N + "|" + CStr(snums) + "|" + CStr(Index) + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
    tempsstr = tempsstr + CStr(W_X_CON) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(SStart_X) + "|" + CStr(SStart_Y) + "|" + CStr(SStart_Z) + "|" + CStr(Bales_Whirl)
    Print #1, tempsstr
    '剩余空间充填
    If S_type = 0 Then '采用深度搜索
        Respace2 CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace2 CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace2 CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
        Respace2 CON_I_N, CON_NS, S_type, Re_X4, Re_Y4, Re_Z4, snums, CStr(Index) + "4", S_Start_X4, S_Start_Y4, S_Start_Z4
    Else '采用浅度搜索
        Respace CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
        Respace CON_I_N, CON_NS, S_type, Re_X4, Re_Y4, Re_Z4, snums, CStr(Index) + "4", S_Start_X4, S_Start_Y4, S_Start_Z4
    End If
End Sub

