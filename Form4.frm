VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "�����嵥"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command12 
      BackColor       =   &H00FFC0C0&
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�鿴װ��ͼ"
      Height          =   540
      Left            =   9660
      TabIndex        =   22
      Top             =   6090
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ�ϲ���ʼ����"
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
         Caption         =   "��X��װ�䷽��"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "ѡ���װ�����"
      Height          =   3585
      Left            =   7455
      TabIndex        =   5
      Top             =   105
      Width           =   4110
      Begin VB.OptionButton Option2 
         Caption         =   "�������"
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
         Caption         =   "ǳ������"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   1785
         TabIndex        =   20
         Top             =   3045
         Width           =   1485
      End
      Begin VB.CheckBox Check2 
         Caption         =   "�������"
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
         Caption         =   "ǳ������"
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
         Caption         =   "�������"
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
         Caption         =   "��������"
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
         Caption         =   "��������"
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
         Caption         =   "�������"
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
         Caption         =   "ʣ��ռ��ֲ��ԣ�"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   2625
         Width           =   3690
      End
      Begin VB.Label Label4 
         Caption         =   "�������ֲ��ԣ�"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1680
         Width           =   3795
      End
      Begin VB.Label Label3 
         Caption         =   "װ�����Ȳ��ԣ�"
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   420
         Width           =   3690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "װ���б�"
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
            Text            =   "ѡ��"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "��(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "��(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "��(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "����(g)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "�Ը�(mm)"
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
            Text            =   "ѡ��"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "��(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "��(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "��(mm)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "����(g)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "����"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "��װ��Ļ��"
         Height          =   330
         Left            =   210
         TabIndex        =   4
         Top             =   3255
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "ѡ���������"
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
'���ר��
'****************************************************************************
Private containers(), bales() As String
Private m_count As Integer
Private dt_x As Boolean      '�Ƿ�ʹ���˻������ռ�
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
'����ʱ�ļ�
Open App.Path + "\temps.txt" For Output As #1
Open App.Path + "\errs.txt" For Output As #2
'ѭ������
pro_index = 1
For i = 1 To CInt(ListView2.ListItems.Count)
    
    container_index = i
    'ѭ�����Ȳ���
    For Each checks1 In Check1
        If checks1.Value = 1 Then
            order_flag = CInt(checks1.Index)
            pro_str2 = " |���Ȳ���:" + checks1.Caption
            'ѭ�����������
            For Each checks2 In Check2
                If checks2.Value = 1 Then
                    If checks2.Index = 0 Then
                        WF_flag = 0 '�������
                    Else
                        WF_flag = 1 'ǳ������
                    End If
                    pro_str3 = " |���������:" + checks2.Caption
                    'ѭ��ʣ��ռ����
                        For Each checks3 In Option2
                            If checks3.Value = True Then
                                If checks3.Index = 0 Then
                                    SR_flag = 0 '�������
                                Else
                                    SR_flag = 1 'ǳ������
                                End If
                                pro_str4 = " |ʣ��ռ����:" + checks3.Caption
                                PRO_STR1 = "PRO_NUM|" + CStr(pro_index) + "|����:" + ListView2.ListItems(i).SubItems(1)
                                Print #1, PRO_STR1 + pro_str2 + pro_str3 + pro_str4
                                '��ʼʵ�ʼ���
                                Get_Container '���м������������
                                Get_Bale
                                OrderBales order_flag '���Ż����б�
                                '����
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
'�ر��ļ�
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
'�����м�����
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
    containers(i, 1) = items.SubItems(1) '����
    containers(i, 2) = items.SubItems(2) '��
    containers(i, 3) = items.SubItems(3) '��
    containers(i, 4) = items.SubItems(4) '��
    containers(i, 5) = CStr(CDbl(items.SubItems(2)) * CDbl(items.SubItems(3)) * CDbl(items.SubItems(4))) '���
    containers(i, 7) = CStr(items.Index)
    i = i + 1
Next
End Function
'��������
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
    bales(i, 1) = items.SubItems(1) '����
    bales(i, 2) = items.SubItems(2) '��
    bales(i, 3) = items.SubItems(3) '��
    bales(i, 4) = items.SubItems(4) '��
    bales(i, 5) = CStr(CDbl(items.SubItems(2)) * CDbl(items.SubItems(3)) * CDbl(items.SubItems(4))) '���
    bales(i, 6) = ""
    bales(i, 7) = items.SubItems(6)
    bales(i, 8) = CStr(items.Index)
    i = i + 1
Next
End Function
'���ȼ����� ���ֻ����װʱ��ѡ����ʲô��׼Ϊ���ȼ�
Private Sub OrderBales(ByVal types As Integer)
Dim tmps(1 To 8)
Select Case types
Case 0 '����Ŵ���
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
Case 1 '�����ų���
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
Case 2 '����Ŵ���
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
Case 3 '�����Ŷ���
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

'С������
Private Function Tests(ByVal S_type As Integer, ByVal C_index As Integer) As Boolean
'S_type ʣ��ռ���������, C_index ���������������е�����
    '================�������===================
    Dim Max_X, Max_Y, Max_Z, WFace_X, WFace_Y, WFace_Z, CON_V, Bales_V, CON_EFF As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '����ĳߴ�
    Dim CON_N, T_Flag_Ok, SNum, TempIndex, TempsIndexs, Bales_N, Bales_Whirl As Integer
    Dim Con_Index_Num, Bales_Index As String
    Dim i As Integer
    Dim Start_X, Start_Y, Start_Z, SMax_X As Double
    '�������������
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3 As Double '��ǰ������ʣ��ռ�ĳߴ�
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3 As Double 'ʣ��ռ��������
    Dim W_Z_N, W_Y_N, W_Z_CON, Can_Count As Integer
    '===========================================
    '================��ʼ������=================
    '�������ߴ�
    Max_X = CDbl(containers(C_index, 2))
    Max_Y = CDbl(containers(C_index, 3))
    Max_Z = CDbl(containers(C_index, 4))
    Con_Index_Num = containers(C_index, 7) '������LISTVIEW2�е�LISTITEM���
    CON_V = 0 '���������
    CON_N = 1 'ʹ�õ���������
    T_Flag_Ok = 0 'װ����ϱ�־ 0 δ��� 1���
    Bales_V = 0 '���������
    '��ȡ���������
    For i = 1 To UBound(bales, 1)
        Bales_V = Bales_V + CDbl(bales(i, 5)) * CInt(bales(i, 7))
    Next i
    '===========================================
    '================��ʼװ�����===============
    Do '����ѭ��
        SNum = 0 '��������
        '��ǰ�������������
        Start_X = 0
        Start_Y = 0
        Start_Z = 0
        dt_x = False
        '���ù����ռ����ߴ���������ߴ�
        WFace_X = Max_X
        WFace_Y = Max_Y
        WFace_Z = Max_Z
        SMax_X = Max_X ' X�����Ͽ��õ����ߴ�
        '������ѭ����ʼ
        Do
            '���ò���
            TempIndex = 0 'ѡ��Ļ����ڻ��������е�����
            '���õ�ǰ��������ÿռ�ߴ�
            WFace_Y = Max_Y
            WFace_Z = Max_Z
            WFace_X = SMax_X
            'ѡ�������Ƚ���װ��
            For i = 1 To UBound(bales, 1) '��������������Ѱ���Է��빤����ռ�Ļ���
                If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '�Ƿ��װ�أ�����ʣ����������0
                    'ѡ�н���װ��Ļ���ĳߴ�
                    Bales_X = CDbl(bales(i, 2))
                    Bales_Y = CDbl(bales(i, 3))
                    Bales_Z = CDbl(bales(i, 4))
                    'ѡ�н���װ��Ļ�������
                    Bales_N = CInt(bales(i, 7))
                    Bales_Index = bales(i, 8) 'ѡ�еĻ�����LISTVIEW3�е�LISTITEM���
                    Bales_Whirl = 0 '���û����Ƿ�ˮƽ��ת��־ 0 δ��ת 1 ��ת90�� Bales_X��Bales_Y����
                    '���������ڳ���
                    '������ﳤ�ȴ��ڹ�������ó��ȵ�С�ڹ�������ó��ȼӻ������ĳ���
                    If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False And hc Then  'New Code
                        dt_x = True '��������ʹ��
                        WFace_X = WFace_X + 100 '���ù�������ó��ȵ��ڹ�������ó��ȼӻ���������
                        TempIndex = i '���û����ڻ��������е��������
                        Exit For '�˳�ѡ�����ѭ��
                    ElseIf Bales_X < WFace_X Then '���ﳤ��С�ڹ�������ó���
                        TempIndex = i
                        Exit For
                    End If
                End If
            Next i
            '���û�к���װ�سߴ�Ļ������ˮƽ��ת
            If TempIndex = 0 Then
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then  '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
                        Bales_Y = CDbl(bales(i, 2))
                        Bales_X = CDbl(bales(i, 3))
                        Bales_Z = CDbl(bales(i, 4))
                        Bales_N = CInt(bales(i, 7))
                        Bales_Index = bales(i, 8)
                        Bales_Whirl = 1
                        '���������ڳ���
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
            '�����û�п�װ�صĻ���
            If TempIndex = 0 Then
                '�ж��Ƿ�װ�����
                TempsIndexs = 0
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 Then '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
                        TempsIndexs = i
                        Exit For
                    End If
                Next i
                If TempsIndexs = 0 Then
                    Tests = True 'װ��ɹ�
                    T_Flag_Ok = 1
                    Exit Do 'װ��ɹ�
                Else
                    '��ǰ����ʣ��ռ��޷����º��ʵ�����,ʹ����һ������
                    Exit Do '�˳�������ѭ��
                End If
            End If
            SNum = SNum + 1 '���ù�������
            

     '��ȷ����Ͽ�װ�������
     W_Y_N = Int(WFace_Y / Bales_Y)
    '��ֱ�����Ͽ�װ�������
    W_Z_N = Int(WFace_Z / Bales_Z)
    '�������װ�����ӵ������Ƿ�С����������
    If W_Y_N * W_Z_N > Bales_N Then
        Can_Count = Bales_N '��װ�ص�������
    Else
        Can_Count = W_Y_N * W_Z_N
    End If
    If (Can_Count Mod W_Y_N) > 0 Then '��Z������װ�ص�����
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
    If (Can_Count Mod W_Y_N) = 0 Then '��ʣ��ռ�2
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
        Else '��ʣ��ռ�3
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
            SMax_X = SMax_X - Bales_X 'ʣ�೤�ȷ�����óߴ�
            '����װ��ͼ
            '��N������|������LISTVIEW������|��������|�ӹ�������|������LISTVIEW������
            '|��������|װ�صĻ�������|X�����װ�ص�����|Y�����װ�ص�����|Z�����װ�ص�����
            '|���X����|���Y����|���Z����|�Ƿ�ˮƽ��ת
            tempsstr = CStr(CON_N) + "|" + Con_Index_Num + "|" + CStr(SNum) + "|" + CStr("0") + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
            tempsstr = tempsstr + CStr(1) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(Start_X) + "|" + CStr(Start_Y) + "|" + CStr(Start_Z) + "|" + CStr(Bales_Whirl)
            Print #1, tempsstr            'Debug.Print "�м���� " + CStr(snum) + " " + CStr(S_type)
            'Debug.Print "ǳ������ ��ӡ������װ���б�"
            'ʣ��ռ����
            If S_type = 0 Then '�����������
                'Debug.Print "ǳ������ ʣ���������"
                Respace2 Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace2 Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace2 Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
            Else '����ǳ������
                'Debug.Print "ǳ������ ʣ��ǳ������"
                Respace Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
            End If

            '��һ�������������
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
                CON_V = CON_V + (containers(C_index, 2) + 100) * containers(C_index, 3) * containers(C_index, 4) '�����˻���������������
            End If
        End If
    Loop
    '����Ч��
    CON_EFF = Bales_V / CON_V
    Print #1, "PRO_EFF=|" + CStr(CON_EFF) + "| CON_N=" + CStr(CON_N)
End Function
'������
Private Function Tests2(ByVal S_type As Integer, ByVal C_index As Integer) As Boolean
'S_type ʣ��ռ���������, C_index ���������������е�����
    '================�������===================
    Dim Max_X, Max_Y, Max_Z, WFace_X, WFace_Y, WFace_Z, CON_V, Bales_V, CON_EFF As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '����ĳߴ�
    Dim CON_N, T_Flag_Ok, SNum, TempIndex, TempsIndexs, Bales_N, Bales_Whirl As Integer
    Dim Con_Index_Num, Bales_Index As String
    Dim i As Integer
    Dim Start_X, Start_Y, Start_Z, SMax_X As Double
    '�������������
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3, Re_X4, Re_Y4, Re_Z4 As Double '��ǰ������ʣ��ռ�ĳߴ�
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3, S_Start_X4, S_Start_Y4, S_Start_Z4 As Double 'ʣ��ռ��������
    Dim W_Z_N, W_Y_N, W_X_N, W_X_CON, Can_Count As Integer
    '===========================================
    '================��ʼ������=================
    '�������ߴ�
    Max_X = CDbl(containers(C_index, 2))
    Max_Y = CDbl(containers(C_index, 3))
    Max_Z = CDbl(containers(C_index, 4))
    Con_Index_Num = containers(C_index, 7) '������LISTVIEW2�е�LISTITEM���
    CON_V = 0 '���������
    CON_N = 1 'ʹ�õ���������
    T_Flag_Ok = 0 'װ����ϱ�־ 0 δ��� 1���
    Bales_V = 0 '���������
    '��ȡ���������
    For i = 1 To UBound(bales, 1)
        Bales_V = Bales_V + CDbl(bales(i, 5)) * CInt(bales(i, 7))
    Next i
    '===========================================
    '================��ʼװ�����===============
    Do '����ѭ��
        SNum = 0 '��������
        '��ǰ�������������
        Start_X = 0
        Start_Y = 0
        Start_Z = 0
        dt_x = False
        '���ù����ռ����ߴ���������ߴ�
        WFace_X = Max_X
        WFace_Y = Max_Y
        WFace_Z = Max_Z
        SMax_X = Max_X ' X�����Ͽ��õ����ߴ�
        '������ѭ����ʼ
        Do
            '���ò���
            TempIndex = 0 'ѡ��Ļ����ڻ��������е�����
            '���õ�ǰ��������ÿռ�ߴ�
            WFace_Y = Max_Y
            WFace_Z = Max_Z
            WFace_X = SMax_X
            'ѡ�������Ƚ���װ��
            For i = 1 To UBound(bales, 1) '��������������Ѱ���Է��빤����ռ�Ļ���
                If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '�Ƿ��װ�أ�����ʣ����������0
                    'ѡ�н���װ��Ļ���ĳߴ�
                    Bales_X = CDbl(bales(i, 2))
                    Bales_Y = CDbl(bales(i, 3))
                    Bales_Z = CDbl(bales(i, 4))
                    'ѡ�н���װ��Ļ�������
                    Bales_N = CInt(bales(i, 7))
                    Bales_Index = bales(i, 8) 'ѡ�еĻ�����LISTVIEW3�е�LISTITEM���
                    Bales_Whirl = 0 '���û����Ƿ�ˮƽ��ת��־ 0 δ��ת 1 ��ת90�� Bales_X��Bales_Y����
                    '���������ڳ���
                    '������ﳤ�ȴ��ڹ�������ó��ȵ�С�ڹ�������ó��ȼӻ������ĳ���
                    If Bales_X > WFace_X And Bales_X < WFace_X + 100 And dt_x = False And hc Then 'New Code
                        dt_x = True '��������ʹ��
                        WFace_X = WFace_X + 100 '���ù�������ó��ȵ��ڹ�������ó��ȼӻ���������
                        TempIndex = i '���û����ڻ��������е��������
                        Exit For '�˳�ѡ�����ѭ��
                    ElseIf Bales_X < WFace_X Then '���ﳤ��С�ڹ�������ó���
                        TempIndex = i
                        Exit For
                    End If
                End If
            Next i
            '���û�к���װ�سߴ�Ļ������ˮƽ��ת
            If TempIndex = 0 Then
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then  '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
                        Bales_Y = CDbl(bales(i, 2))
                        Bales_X = CDbl(bales(i, 3))
                        Bales_Z = CDbl(bales(i, 4))
                        Bales_N = CInt(bales(i, 7))
                        Bales_Index = bales(i, 8)
                        Bales_Whirl = 1
                        '���������ڳ���
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
            '�����û�п�װ�صĻ���
            If TempIndex = 0 Then
                '�ж��Ƿ�װ�����
                TempsIndexs = 0
                For i = 1 To UBound(bales, 1)
                    If CInt(bales(i, 7)) > 0 Then '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
                        TempsIndexs = i
                        Exit For
                    End If
                Next i
                If TempsIndexs = 0 Then
                    Tests2 = True 'װ��ɹ�
                    T_Flag_Ok = 1
                    Exit Do 'װ��ɹ�
                Else
                    '��ǰ����ʣ��ռ��޷����º��ʵ�����,ʹ����һ������
                    Exit Do '�˳�������ѭ��
                End If
            End If
            SNum = SNum + 1 '���ù�������
            '��ȷ����Ͽ�װ�������
            W_Y_N = Int(WFace_Y / Bales_Y)
            '��ֱ�����Ͽ�װ�������
            W_Z_N = Int(WFace_Z / Bales_Z)
            '���ȷ����Ͽ�װ�������
            W_X_N = Int(WFace_X / Bales_X)
            '�������װ�����ӵ������Ƿ�С����������
            If W_Y_N * W_Z_N * W_X_N > Bales_N Then
                Can_Count = Bales_N
            Else
                Can_Count = W_Y_N * W_Z_N * W_X_N
            End If
            '���1
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
                'ʣ��ռ��������
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
            SMax_X = SMax_X - Bales_X * W_X_CON 'ʣ�೤�ȷ�����óߴ�
            tempsstr = CStr(CON_N) + "|" + Con_Index_Num + "|" + CStr(SNum) + "|" + CStr("0") + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
            tempsstr = tempsstr + CStr(W_X_CON) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(Start_X) + "|" + CStr(Start_Y) + "|" + CStr(Start_Z) + "|" + CStr(Bales_Whirl)
            Print #1, tempsstr
            'ʣ��ռ����
            If S_type = 0 Then '�����������
                Respace2 Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace2 Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace2 Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
                Respace2 Con_Index_Num, CON_N, S_type, Re_X4, Re_Y4, Re_Z4, SNum, 4, S_Start_X4, S_Start_Y4, S_Start_Z4
            Else '����ǳ������
                Respace Con_Index_Num, CON_N, S_type, Re_X1, Re_Y1, Re_Z1, SNum, 1, S_Start_X1, S_Start_Y1, S_Start_Z1
                Respace Con_Index_Num, CON_N, S_type, Re_X2, Re_Y2, Re_Z2, SNum, 2, S_Start_X2, S_Start_Y2, S_Start_Z2
                Respace Con_Index_Num, CON_N, S_type, Re_X3, Re_Y3, Re_Z3, SNum, 3, S_Start_X3, S_Start_Y3, S_Start_Z3
                Respace Con_Index_Num, CON_N, S_type, Re_X4, Re_Y4, Re_Z4, SNum, 4, S_Start_X4, S_Start_Y4, S_Start_Z4
            End If
            '��һ�������������
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
                CON_V = CON_V + (containers(C_index, 2) + 100) * containers(C_index, 3) * containers(C_index, 4) '�����˻���������������
            End If
        End If
        DoEvents
    Loop
    '����Ч��
    CON_EFF = Bales_V / CON_V
    Print #1, "PRO_EFF=|" + CStr(CON_EFF) + "| CON_N=" + CStr(CON_N)
End Function
'����ʣ��ռ�
Private Sub Respace(ByVal CON_I_N As String, ByVal CON_NS As Integer, ByVal S_type As Integer, x_max, y_max, z_max, snums, Index, SStart_X, SStart_Y, SStart_Z)
    '================�������===================
    Dim WFace_X, WFace_Y, WFace_Z As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '����ĳߴ�
    Dim TempIndex, Bales_N, Bales_Whirl As Integer
    Dim Bales_Index As String
    Dim i As Integer
    '�������������
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3 As Double '��ǰ������ʣ��ռ�ĳߴ�
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3 As Double 'ʣ��ռ��������
    Dim W_Z_N, W_Y_N, W_Z_CON, Can_Count As Integer
    '===========================================
    TempIndex = 0
    WFace_Y = y_max
    WFace_Z = z_max
    WFace_X = x_max
    'ѡ�������Ƚ���װ��
    
    For i = 1 To UBound(bales, 1)
        If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
            Bales_X = CDbl(bales(i, 2))
            Bales_Y = CDbl(bales(i, 3))
            Bales_Z = CDbl(bales(i, 4))
            Bales_N = CInt(bales(i, 7))
            Bales_Index = bales(i, 8)
            Bales_Whirl = 0
            '���������ڳ���
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
    '���û�к���װ�سߴ�Ļ������ˮƽ��ת
    If TempIndex = 0 Then
        For i = 1 To UBound(bales, 1)
            If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then   '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
                Bales_Y = CDbl(bales(i, 2))
                Bales_X = CDbl(bales(i, 3))
                Bales_Z = CDbl(bales(i, 4))
                Bales_N = CInt(bales(i, 7))
                Bales_Index = bales(i, 8)
                Bales_Whirl = 1
                '���������ڳ���
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
    '��ȷ����Ͽ�װ�������
     W_Y_N = Int(WFace_Y / Bales_Y)
    '��ֱ�����Ͽ�װ�������
    W_Z_N = Int(WFace_Z / Bales_Z)
    '�������װ�����ӵ������Ƿ�С����������
    If W_Y_N * W_Z_N > Bales_N Then
        Can_Count = Bales_N '��װ�ص�������
    Else
        Can_Count = W_Y_N * W_Z_N
    End If
    If (Can_Count Mod W_Y_N) > 0 Then '��Z������װ�ص�����
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
    If (Can_Count Mod W_Y_N) = 0 Then '��ʣ��ռ�2
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
        Else '��ʣ��ռ�3
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
    '����װ��ͼ
    tempsstr = CStr(CON_NS) + "|" + CON_I_N + "|" + CStr(snums) + "|" + CStr(Index) + "|" + Bales_Index + "|" + CStr(bales(TempIndex, 1)) + "|" + CStr(Can_Count) + "|"
    tempsstr = tempsstr + CStr(1) + "|" + CStr(W_Y_N) + "|" + CStr(W_Z_N) + "|" + CStr(SStart_X) + "|" + CStr(SStart_Y) + "|" + CStr(SStart_Z) + "|" + CStr(Bales_Whirl)
    Print #1, tempsstr
    'ʣ��ռ����
    If S_type = 0 Then '�����������
        Respace2 CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace2 CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace2 CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
    Else '����ǳ������
        Respace CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
    End If
End Sub
Private Sub Respace2(ByVal CON_I_N As String, ByVal CON_NS As Integer, ByVal S_type As Integer, x_max, y_max, z_max, snums, Index, SStart_X, SStart_Y, SStart_Z)
    '================�������===================
    Dim WFace_X, WFace_Y, WFace_Z As Double
    Dim Bales_X, Bales_Y, Bales_Z As Double '����ĳߴ�
    Dim TempIndex, Bales_N, Bales_Whirl As Integer
    Dim Bales_Index As String
    Dim i As Integer
    '�������������
    Dim Re_X1, Re_Y1, Re_Z1, Re_X2, Re_Y2, Re_Z2, Re_X3, Re_Y3, Re_Z3, Re_X4, Re_Y4, Re_Z4 As Double '��ǰ������ʣ��ռ�ĳߴ�
    Dim S_Start_X1, S_Start_Y1, S_Start_Z1, S_Start_X2, S_Start_Y2, S_Start_Z2, S_Start_X3, S_Start_Y3, S_Start_Z3, S_Start_X4, S_Start_Y4, S_Start_Z4 As Double 'ʣ��ռ��������
    Dim W_Z_N, W_Y_N, W_X_N, W_X_CON, Can_Count As Integer
    '===========================================
    TempIndex = 0
    WFace_Y = y_max
    WFace_Z = z_max
    WFace_X = x_max
    'ѡ�������Ƚ���װ��
    For i = 1 To UBound(bales, 1)
        If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 3)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
            Bales_X = CDbl(bales(i, 2))
            Bales_Y = CDbl(bales(i, 3))
            Bales_Z = CDbl(bales(i, 4))
            Bales_N = CInt(bales(i, 7))
            Bales_Index = bales(i, 8)
            Bales_Whirl = 0
            '���������ڳ���
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
    '���û�к���װ�سߴ�Ļ������ˮƽ��ת
    If TempIndex = 0 Then
        For i = 1 To UBound(bales, 1)
            If CInt(bales(i, 7)) > 0 And CDbl(bales(i, 2)) < WFace_Y And CDbl(bales(i, 4)) < WFace_Z Then   '�Ƿ�δװ�أ�����ʣ����������0 Bales(i, 0) = 0 And
                Bales_Y = CDbl(bales(i, 2))
                Bales_X = CDbl(bales(i, 3))
                Bales_Z = CDbl(bales(i, 4))
                Bales_N = CInt(bales(i, 7))
                Bales_Index = bales(i, 8)
                Bales_Whirl = 1
                '���������ڳ���
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
    '��ȷ����Ͽ�װ�������
    W_Y_N = Int(WFace_Y / Bales_Y)
    '��ֱ�����Ͽ�װ�������
    W_Z_N = Int(WFace_Z / Bales_Z)
    '���ȷ����Ͽ�װ�������
    W_X_N = Int(WFace_X / Bales_X)
    '�������װ�����ӵ������Ƿ�С����������
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
        'ʣ��ռ��������
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
    'ʣ��ռ����
    If S_type = 0 Then '�����������
        Respace2 CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace2 CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace2 CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
        Respace2 CON_I_N, CON_NS, S_type, Re_X4, Re_Y4, Re_Z4, snums, CStr(Index) + "4", S_Start_X4, S_Start_Y4, S_Start_Z4
    Else '����ǳ������
        Respace CON_I_N, CON_NS, S_type, Re_X1, Re_Y1, Re_Z1, snums, CStr(Index) + "1", S_Start_X1, S_Start_Y1, S_Start_Z1
        Respace CON_I_N, CON_NS, S_type, Re_X2, Re_Y2, Re_Z2, snums, CStr(Index) + "2", S_Start_X2, S_Start_Y2, S_Start_Z2
        Respace CON_I_N, CON_NS, S_type, Re_X3, Re_Y3, Re_Z3, snums, CStr(Index) + "3", S_Start_X3, S_Start_Y3, S_Start_Z3
        Respace CON_I_N, CON_NS, S_type, Re_X4, Re_Y4, Re_Z4, snums, CStr(Index) + "4", S_Start_X4, S_Start_Y4, S_Start_Z4
    End If
End Sub

