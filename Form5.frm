VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   BackColor       =   &H8000000D&
   Caption         =   "装箱图示例(参考)"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   7320
   ScaleWidth      =   11310
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   3060
      TabIndex        =   1
      Top             =   1575
      Width           =   3060
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4005
         Left            =   0
         TabIndex        =   3
         Top             =   105
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   7064
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         SingleSel       =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   0
         Picture         =   "Form5.frx":0ECA
         Top             =   4920
         Width           =   720
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5430
      Left            =   3045
      ScaleHeight     =   5400
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   1680
      Width           =   8265
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   630
         Left            =   4305
         TabIndex        =   4
         Top             =   4725
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2325
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   4101
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "方案"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "效率"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "容器"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "策略"
         Object.Width           =   14111
      EndProperty
   End
   Begin VB.Menu popmun 
      Caption         =   "popmun"
      Visible         =   0   'False
      Begin VB.Menu popline4 
         Caption         =   "-"
      End
      Begin VB.Menu popclear 
         Caption         =   "擦除图形"
      End
      Begin VB.Menu popline1 
         Caption         =   "-"
      End
      Begin VB.Menu popsave 
         Caption         =   "保存图形"
      End
      Begin VB.Menu popline3 
         Caption         =   "-"
      End
      Begin VB.Menu popprint 
         Caption         =   "打印图形"
      End
      Begin VB.Menu popline2 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************************
'申皇专用
'****************************************************************************

Private Sub Form_Load()
ListView1.Height = 2400
ListView1.Top = 0
ListView1.Left = 0
ListView1.Width = Me.ScaleWidth
Picture1.Left = 3060 + 50
Picture1.Top = ListView1.Height + 50
Picture1.Height = Me.ScaleHeight - 2400 - 50
Picture2.Left = 0
Picture2.Top = ListView1.Height + 50
Picture2.Height = Me.ScaleHeight - 2400 - 50
TreeView1.Height = Picture2.ScaleHeight
TreeView1.Top = 0
TreeView1.Left = 0
Picture2.Width = 3060
Picture1.Width = Me.ScaleWidth - 3060 - 50
Picture1.ScaleMode = 1
Label1.Width = Picture1.ScaleWidth - 400
Label1.Left = 200
Label1.Top = Picture1.ScaleHeight - 620
init_pic



End Sub
Private Sub init_treeview(ByVal se As String)
pro_name = ""
root_index = 0
con_num_i = 0
st_end = Split(se, "|")
TreeView1.Nodes.Clear
For i = CInt(st_end(0)) To CInt(st_end(1))
        textline = ss(i)
        If Mid(textline, 1, 7) = "PRO_NUM" Then
            pic_label = Replace(textline, "|", " ")
        ElseIf Mid(textline, 1, 7) = "PRO_EFF" Then
            temps = Split(textline, "|")
            pic_label = pic_label + " 方案效率：" + CStr(Format(temps(1), "0.00%"))
        Else
            temps = Split(textline, "|")
            
            If con_num_i <> CInt(temps(0)) Then
                pro_name = "PRO_" & temps(0)
                con_num_i = CInt(temps(0))
                Set nodX = TreeView1.Nodes.Add(, , pro_name, "容器：" & temps(0))
                root_index = nodX.Index
            End If
            If CInt(temps(3)) > 0 Then
                Texts1 = pro_name & "WC" & temps(2) & "_" & temps(3)
                Texts2 = "工作面：" & temps(2) & " 子工作面：" & temps(3)
            Else
                Texts1 = pro_name & "W" & temps(2) & "_" & temps(3)
                Texts2 = "工作面：" & temps(2) & " 子工作面：0"
            End If
        
            Set nodX = TreeView1.Nodes.Add(pro_name, tvwChild, Texts1, Texts2)
            nodX.Tag = textline
        End If
Next
For Each nods In TreeView1.Nodes
    If nods.Children > 0 Then
        nods.Tag = pic_label
    End If
Next
nodX.EnsureVisible
End Sub
Private Sub init_pic()
Open App.Path + "\temps.txt" For Input As #2
Seek #2, 1

i = 0
Dim nodX As Node
            '第N个容器|容器在LISTVIEW中索引|工作面编号|子工作面编号|货物再LISTVIEW中索引
            '|货物名称|装载的货物数量|X方向可装载的数量|Y方向可装载的数量|Z方向可装载的数量
            '|起点X坐标|起点Y坐标|起点Z坐标|是否水平旋转
    Do While Not EOF(2)   ' 循环至文件尾。
        
        i = i + 1
        ReDim Preserve ss(1 To i)
        Line Input #2, textline   ' 读入一行数据。
        If Mid(textline, 1, 7) = "PRO_NUM" Then
            temps = Split(textline, "|")
                Set itmx = ListView1.ListItems.Add(, , temps(1))
                itmx.Tag = CStr(i)
                itmx.SubItems(2) = temps(2)
                itmx.SubItems(3) = temps(3) + " " + temps(4) + " " + temps(5)
                ss(i) = textline
        ElseIf Mid(textline, 1, 7) = "PRO_EFF" Then
                temps = Split(textline, "|")
                itmx.SubItems(1) = CStr(Format(temps(1), "0.00%"))
                itmx.Tag = itmx.Tag + "|" + CStr(i)
                ss(i) = textline
        Else
                      
                ss(i) = textline

        End If
        DoEvents
    Loop
    'paintpic
    'For i = 1 To UBound(ss, 2)
        'Paintboxs (i - 1)
    'Next i
Close
End Sub
Private Sub Form_Resize()
ListView1.Height = 2400
ListView1.Top = 0
ListView1.Left = 0
ListView1.Width = Me.ScaleWidth
Picture1.Left = 3060 + 50
Picture1.Top = ListView1.Height + 50
Picture1.Height = Me.ScaleHeight - 2400 - 50
Picture2.Left = 0
Picture2.Top = ListView1.Height + 50
Picture2.Height = Me.ScaleHeight - 2400 - 50
TreeView1.Height = Picture2.ScaleHeight
TreeView1.Top = 0
TreeView1.Left = 0
Picture2.Width = 3060
Picture1.Width = Me.ScaleWidth - 3060 - 50
Picture1.ScaleMode = 1
Label1.Width = Picture1.ScaleWidth - 400
Label1.Left = 200
Label1.Top = Picture1.ScaleHeight - 620

End Sub
Private Sub ListView1_Click()
init_treeview ListView1.SelectedItem.Tag
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu popmun
End If
End Sub
Private Sub popclear_Click()
Picture1.Cls
Label1.Caption = ""
End Sub

Private Sub TreeView1_Click()
'MsgBox TreeView1.SelectedItem.Tag
If TreeView1.SelectedItem.Children > 0 Then

Label1.Caption = TreeView1.SelectedItem.Tag
Else
Label1.Caption = TreeView1.SelectedItem.Parent.Tag + "  " + TreeView1.SelectedItem.Text



End If
End Sub
Private Sub TreeView1_DblClick()
If TreeView1.SelectedItem.Children > 0 Then
temps = Split(TreeView1.SelectedItem.Child.Tag, "|")
Picture1.Cls
paintpic temps(1)


For Each nots In TreeView1.Nodes
If nots.Children = 0 Then
If nots.Parent.Index = TreeView1.SelectedItem.Index Then
    Paintboxs nots.Tag
End If
End If
Next
Label1.Caption = TreeView1.SelectedItem.Tag
Else
Label1.Caption = TreeView1.SelectedItem.Parent.Tag + "  " + TreeView1.SelectedItem.Text
temps = Split(TreeView1.SelectedItem.Tag, "|")

paintpic temps(1)
Paintboxs TreeView1.SelectedItem.Tag
End If

End Sub

'画箱子装载图
Private Sub Paintboxs(ByVal paintstr As String)
    Index = Index + 1
    Picture1.ForeColor = RGB(255, 0, 0)
    '查找箱子尺寸
    Debug.Print paintstr
    
    temps = Split(paintstr, "|")
    Debug.Print temps(4)
            '0第N个容器|1容器在LISTVIEW中索引|2工作面编号|3子工作面编号|4货物再LISTVIEW中索引
            '|5货物名称|6装载的货物数量|7X方向可装载的数量|7Y方向可装载的数量|9Z方向可装载的数量
            '|10起点X坐标|11起点Y坐标|12起点Z坐标|13是否水平旋转
         
            Bales_X = CDbl(Form4.ListView3.ListItems(CInt(temps(4))).SubItems(2))
            Bales_Y = CDbl(Form4.ListView3.ListItems(CInt(temps(4))).SubItems(3))
            Bales_Z = CDbl(Form4.ListView3.ListItems(CInt(temps(4))).SubItems(4))
            
            dt_color = Int(155 / Form4.ListView3.ListItems.Count) * CInt(temps(4))
            dt_color1 = Int(255 / Form4.ListView3.ListItems.Count) * CInt(temps(4))
            dt_color2 = Int(200 / Form4.ListView3.ListItems.Count) * CInt(temps(4))
    Debug.Print Bales_X
    Debug.Print Bales_Y
    Debug.Print Bales_Z
    If CInt(temps(13)) = 1 Then
        temp = Bales_X
        Bales_X = Bales_Y
        Bales_Y = temp
    End If
    '画箱子
    box_count = 0
    For i = 1 To CInt(temps(7))
        For j = 1 To CInt(temps(9))
            For k = 1 To CInt(temps(8))
                box_count = box_count + 1
                Dim tips(1 To 8, 1 To 2)
                s_x = CInt(temps(10)) - Int(CInt(temps(11)) / 1.414)
                s_y = Int(CInt(temps(11)) / 1.414) - CInt(temps(12))
                tips(2, 1) = s_x + Bales_X * i - Int(Bales_Y / 1.414) * (k - 1)
                tips(2, 2) = s_y + Int(Bales_Y / 1.414) * (k - 1) - Bales_Z * (j - 1)
                tips(3, 1) = s_x + Bales_X * i - Int(Bales_Y / 1.414) * k
                tips(3, 2) = s_y + Int(Bales_Y / 1.414) * k - Bales_Z * (j - 1)
                tips(4, 1) = s_x + Bales_X * (i - 1) - Int(Bales_Y / 1.414) * k
                tips(4, 2) = s_y + Int(Bales_Y / 1.414) * k - Bales_Z * (j - 1)
                tips(5, 1) = s_x + Bales_X * (i - 1) - Int(Bales_Y / 1.414) * (k - 1)
                tips(5, 2) = s_y + Int(Bales_Y / 1.414) * (k - 1) - Bales_Z * j
                tips(6, 1) = s_x + Bales_X * i - Int(Bales_Y / 1.414) * (k - 1)
                tips(6, 2) = s_y + Int(Bales_Y / 1.414) * (k - 1) - Bales_Z * j
                tips(7, 1) = s_x + Bales_X * i - Int(Bales_Y / 1.414) * k
                tips(7, 2) = s_y + Int(Bales_Y / 1.414) * k - Bales_Z * j
                tips(8, 1) = s_x + Bales_X * (i - 1) - Int(Bales_Y / 1.414) * k
                tips(8, 2) = s_y + Int(Bales_Y / 1.414) * k - Bales_Z * j



                For dt_i = 1 To Int(Bales_Y / 1.414)
                Picture1.Line (tips(5, 1) - dt_i, tips(5, 2) + dt_i)-(tips(6, 1) - dt_i, tips(6, 2) + dt_i), RGB(0, dt_color1, dt_color2)
                Next dt_i

                For dt_i = 1 To Int(Bales_Y / 1.414)
                Picture1.Line (tips(6, 1) - dt_i, tips(6, 2) + dt_i)-(tips(2, 1) - dt_i, tips(2, 2) + dt_i), RGB(0, dt_color1, dt_color2)
                Next dt_i                'Picture1.Line (tips(7, 1), tips(7, 2))-(tips(3, 1), tips(3, 2))
                'Picture1.Line (tips(8, 1), tips(8, 2))-(tips(4, 1), tips(4, 2))
                Picture1.Line (tips(8, 1), tips(8, 2))-(tips(3, 1), tips(3, 2)), RGB(0, dt_color1, dt_color2), BF
                Picture1.Line (tips(2, 1), tips(2, 2))-(tips(3, 1), tips(3, 2))
                'Picture1.Line (tips(3, 1), tips(3, 2))-(tips(4, 1), tips(4, 2))
                
                Picture1.Line (tips(5, 1), tips(5, 2))-(tips(6, 1), tips(6, 2))
                 Picture1.Line (tips(6, 1), tips(6, 2))-(tips(7, 1), tips(7, 2))
                'Picture1.Line (tips(7, 1), tips(7, 2))-(tips(8, 1), tips(8, 2))
                Picture1.Line (tips(8, 1), tips(8, 2))-(tips(5, 1), tips(5, 2))
                Picture1.Line (tips(6, 1), tips(6, 2))-(tips(2, 1), tips(2, 2))
                Picture1.Line (tips(8, 1), tips(8, 2))-(tips(3, 1), tips(3, 2)), , B
                If box_count = CInt(temps(6)) Then
                    Exit Sub
                End If
                DoEvents
            Next k
        Next j
    Next i
End Sub
'画图
Private Sub paintpic(ByVal con_index As Integer)
    Picture1.ForeColor = RGB(0, 0, 0)

            con_x = CDbl(Form4.ListView2.ListItems(con_index).SubItems(2))
            con_y = CDbl(Form4.ListView2.ListItems(con_index).SubItems(3))
            con_z = CDbl(Form4.ListView2.ListItems(con_index).SubItems(4))
    Picture1.ScaleMode = 1
    picture1_sc = Picture1.ScaleWidth / Picture1.ScaleHeight
    Picture1.ScaleMode = 0
    
    Picture1.ScaleWidth = 17000   ' 设置宽度的单位值。
    Picture1.ScaleHeight = Picture1.ScaleWidth / picture1_sc ' 设置高度的单位值。
    Picture1.ScaleTop = -6000   ' 顶部设置刻度。
    Picture1.ScaleLeft = -4000   ' 左部设置刻度。
    Picture1.Cls
    '画集装箱
    Dim tips(1 To 8, 1 To 2)
    tips(1, 1) = 0
    tips(1, 2) = 0
    tips(2, 1) = con_x
    tips(2, 2) = 0
    tips(3, 1) = con_x - Int(con_y / 1.414)
    tips(3, 2) = Int(con_y / 1.414)
    tips(4, 1) = -Int(con_y / 1.414)
    tips(4, 2) = Int(con_y / 1.414)
    tips(5, 1) = 0
    tips(5, 2) = -con_z
    tips(6, 1) = con_x
    tips(6, 2) = -con_z
    tips(7, 1) = con_x - Int(con_y / 1.414)
    tips(7, 2) = Int(con_y / 1.414) - con_z
    tips(8, 1) = -Int(con_y / 1.414)
    tips(8, 2) = Int(con_y / 1.414) - con_z
    Picture1.DrawStyle = 1
    Picture1.Line (tips(1, 1), tips(1, 2))-(tips(2, 1), tips(2, 2))
    Picture1.DrawStyle = 0
    Picture1.Line (tips(2, 1), tips(2, 2))-(tips(3, 1), tips(3, 2))
    Picture1.Line (tips(3, 1), tips(3, 2))-(tips(4, 1), tips(4, 2))
    Picture1.DrawStyle = 1
    Picture1.Line (tips(4, 1), tips(4, 2))-(tips(1, 1), tips(1, 2))
    Picture1.DrawStyle = 0
    Picture1.Line (tips(5, 1), tips(5, 2))-(tips(6, 1), tips(6, 2))
    Picture1.Line (tips(6, 1), tips(6, 2))-(tips(7, 1), tips(7, 2))
    Picture1.Line (tips(7, 1), tips(7, 2))-(tips(8, 1), tips(8, 2))
    Picture1.Line (tips(8, 1), tips(8, 2))-(tips(5, 1), tips(5, 2))
    Picture1.DrawStyle = 1
    Picture1.Line (tips(5, 1), tips(5, 2))-(tips(1, 1), tips(1, 2))
    Picture1.DrawStyle = 0
    Picture1.Line (tips(6, 1), tips(6, 2))-(tips(2, 1), tips(2, 2))
    Picture1.Line (tips(7, 1), tips(7, 2))-(tips(3, 1), tips(3, 2))
    Picture1.Line (tips(8, 1), tips(8, 2))-(tips(4, 1), tips(4, 2))
    '打印坐标方向
    Picture1.DrawMode = 13
    Picture1.DrawStyle = 0
    Picture1.PSet (tips(2, 1) + 30, tips(2, 2) - 160), vbWhite
    Picture1.Print "—X"
    Picture1.PSet (tips(5, 1) - 70, tips(5, 2) - 600), vbWhite
    Picture1.Print "Z"
    Picture1.PSet (tips(5, 1) - 100, tips(5, 2) - 300), vbWhite
    Picture1.Print "|"
    Picture1.PSet (tips(4, 1) - 110, tips(4, 2) - 20), vbWhite
    Picture1.Print "/Y"
    
End Sub



