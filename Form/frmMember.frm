VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMember 
   Caption         =   "会员管理"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   ControlBox      =   0   'False
   Icon            =   "frmMember.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   45
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   452
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   45
      Width           =   6780
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "会员管理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   600
         TabIndex        =   18
         Top             =   120
         Width           =   900
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   60
         Top             =   -15
         Width           =   480
      End
   End
   Begin SuperMarket.XPButton cmdClear 
      Height          =   345
      Left            =   3510
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "清理(&C)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin SuperMarket.FTextBox txtCard 
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Top             =   660
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "宋体"
      FontSize        =   9
   End
   Begin MSComctlLib.ListView List1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "会员ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "会员卡号"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "消费金额"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "注册日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "状态"
         Object.Width           =   1323
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMember.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMember.frx":05A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SuperMarket.XPButton cmdDel 
      Height          =   345
      Left            =   2310
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "删除(&D)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin SuperMarket.XPButton cmdEdit 
      Height          =   345
      Left            =   1110
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "修改(&E)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin SuperMarket.XPButton cmdAdd 
      Default         =   -1  'True
      Height          =   345
      Left            =   3510
      TabIndex        =   1
      Top             =   645
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      Caption         =   "添加/查找"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame freItem 
      Height          =   2370
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   4380
      Begin SuperMarket.XPButton cmdToday 
         Height          =   345
         Left            =   3285
         TabIndex        =   15
         Top             =   1305
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         Caption         =   "今天"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin SuperMarket.FTextBox txtDate 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   1320
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   9
         AutoSelAll      =   -1  'True
      End
      Begin SuperMarket.FTextBox txtCard2 
         Height          =   300
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   9
         AutoSelAll      =   -1  'True
      End
      Begin SuperMarket.FTextBox txtCost 
         Height          =   300
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "宋体"
         FontSize        =   9
         AutoSelAll      =   -1  'True
         isNumber        =   -1  'True
         MaxLength       =   10
      End
      Begin SuperMarket.XPButton cmdExit 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   2940
         TabIndex        =   9
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "取消"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin SuperMarket.XPButton cmdOK 
         Height          =   345
         Left            =   1740
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "确定"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "消费金额："
         Height          =   180
         Left            =   360
         TabIndex        =   13
         Top             =   915
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "会员卡号："
         Height          =   180
         Left            =   360
         TabIndex        =   12
         Top             =   435
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记日期："
         Height          =   180
         Left            =   360
         TabIndex        =   11
         Top             =   1380
         Width           =   900
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号："
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   540
   End
End
Attribute VB_Name = "frmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'超市销售系统
'程序开发：lc_mtt
'CSDN博客：http://blog.csdn.net/lc_mtt/
'个人主页：http://www.3lsoft.com
'邮箱：3lsoft@163.com
'注：此代码禁止用于商业用途。有修改者发我一份，谢谢！
'---------------- 开源世界，你我更进步 ----------------

Private Sub cmdAdd_Click()
On Error GoTo aaaa
    txtCard.Text = Trim(txtCard.Text)
    If txtCard.Text = "" Then txtCard.SetFocus: Exit Sub
    Dim i As Long
    For i = 1 To List1.ListItems.Count
        If StrComp(txtCard.Text, List1.ListItems(i).SubItems(1), 1) = 0 Then
            List1.ListItems(i).Selected = True
            SetSB 2, "找到会员卡 " & txtCard.Text & " ."
            txtCard.Text = ""
            txtCard.SetFocus
            Exit Sub
        End If
    Next
    cnMain.Execute "insert [Member] values('" & txtCard.Text & "',0,'" & FormatDate(Date) & "')"
    Dim Item As ListItem
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "Select TOP 1 * From [Member] order by MemberID Desc", cnMain, 1, 1
    b = CheckOutDate(CDate(rs("RegDate")))
    Set Item = List1.ListItems.Add(1, , rs("MemberID"), , 1)
    Item.SubItems(1) = rs("MemberCard")
    Item.SubItems(2) = rs("TotalCost")
    Item.SubItems(3) = rs("RegDate")
    Item.SubItems(4) = "正常"
    Item.Selected = True
    SetSB 2, "已添加会员卡 " & txtCard.Text & " ."
    txtCard.Text = ""
    txtCard.SetFocus
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
    txtCard.SetFocus
End Sub

Public Function FormatDate(ByVal d As Date) As String
    FormatDate = Format(d, "yyyy-mm-dd")
End Function

Private Sub cmdClear_Click()
On Error GoTo aaaa
    Dim i As Long, j As Long, k As Long
    j = List1.ListItems.Count
    If j <= 0 Then
        MsgBox "会员列表为空！", vbInformation
        txtCard.SetFocus
        Exit Sub
    End If
    If MsgBox("这个操作会清理所有的过期会员，请问继续吗？", vbOKCancel + vbExclamation + vbDefaultButton2) = vbCancel Then
        txtCard.SetFocus
        Exit Sub
    End If
    For i = j To 1 Step -1
        If List1.ListItems(i).SmallIcon = 2 Then
            cnMain.Execute "Delete From [Member] Where MemberCard='" & List1.ListItems(i).SubItems(1) & "'"
            List1.ListItems.Remove i
            k = k + 1
        End If
    Next
    MsgBox "清理过程顺利完成，请看以下统计数据：" & vbCrLf & vbCrLf & "原来会员个数： " & j & vbCrLf & "过期会员个数： " & k & vbCrLf & "现在会员个数： " & List1.ListItems.Count, vbInformation
    SetSB 2, "清理过程顺利完成."
    txtCard.SetFocus
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
    LoadMemberList
    txtCard.SetFocus
End Sub

Private Sub cmdDel_Click()
On Error GoTo aaaa
    Dim Item As ListItem
    Set Item = List1.SelectedItem
    If MsgBox("确定删除会员 " & Item.SubItems(1) & " 吗", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
    cnMain.Execute "Delete From [Member] Where MemberCard='" & Item.SubItems(1) & "'"
    SetSB 2, "删除会员卡 " & Item.SubItems(1) & " 成功."
    List1.ListItems.Remove Item.Index
    txtCard.SetFocus
Exit Sub
aaaa:
    If Err.Number <> 91 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdEdit_Click()
On Error GoTo aaaa
    Dim Item As ListItem
    Set Item = List1.SelectedItem
    txtCard2.Text = Item.SubItems(1)
    txtCard2.Tag = Item.SubItems(1)
    txtCost.Text = Item.SubItems(2)
    txtCost.Tag = Item.SubItems(2)
    txtDate.Text = Item.SubItems(3)
    txtDate.Tag = Item.SubItems(3)
    ShowItemFrame True
    txtCard2.SetFocus
aaaa:
End Sub

Private Sub cmdExit_Click()
    ShowItemFrame False
    txtCard.SetFocus
End Sub

Private Sub cmdOK_Click()
On Error GoTo aaaa
    txtCard2.Text = Trim(txtCard2.Text)
    If txtCard2.Text = "" Then
        MsgBox "必须填写会员卡号。", vbInformation
        txtCard2.SetFocus
        Exit Sub
    End If
    cnMain.Execute "UPDATE [Member] SET MemberCard='" & txtCard2.Text & "',TotalCost=" & txtCost.Text & ",RegDate='" & txtDate.Text & "' Where MemberCard='" & txtCard2.Tag & "'"
    Dim Item As ListItem, b As Boolean
    b = CheckOutDate(CDate(txtDate.Text))
    Set Item = List1.SelectedItem
    Item.SmallIcon = IIf(b = False, 1, 2)
    Item.SubItems(1) = txtCard2.Text
    Item.SubItems(2) = txtCost.Text
    Item.SubItems(3) = txtDate.Text
    Item.SubItems(4) = IIf(b = False, "正常", "过期")
    SetSB 2, "修改会员卡 " & txtCard2.Text & " 成功."
    cmdExit_Click
Exit Sub
aaaa:
    MsgBox "操作失败，可能是该会员卡号已经存在！", vbCritical
End Sub

Private Sub cmdToday_Click()
    txtDate.Text = FormatDate(Date)
End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    imgIcon.Picture = frmMain.cmdLeft(5).Picture
    '读取会员数据列表
    LoadMemberList
End Sub

'读取会员数据列表
Public Sub LoadMemberList()
    Dim Item As ListItem, b As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    List1.ListItems.Clear
    rs.Open "Select * From [Member] order by MemberID Desc", cnMain, 1, 1
    Do Until rs.EOF
        b = CheckOutDate(CDate(rs("RegDate")))
        Set Item = List1.ListItems.Add(, , rs("MemberID"), , IIf(b = False, 1, 2))
        Item.SubItems(1) = rs("MemberCard")
        Item.SubItems(2) = rs("TotalCost")
        Item.SubItems(3) = rs("RegDate")
        Item.SubItems(4) = IIf(b = False, "正常", "过期")
        rs.MoveNext
    Loop
    SetSB 2, "共 " & rs.RecordCount & " 条会员记录."
End Sub

Public Function CheckOutDate(ByVal d As Date) As Boolean
    Dim j1 As Long, j2 As Long, j3 As Long
    j1 = Year(Date) - Year(d)
    j2 = Month(Date) - Month(d)
    j3 = Day(Date) - Day(d)
    If j1 > 1 Then
        CheckOutDate = True
    Else
        CheckOutDate = (j1 + j2 + j3 > 0)
    End If
End Function

Public Sub ShowItemFrame(ByVal b As Boolean)
    List1.Visible = Not b
    freItem.Visible = b
    cmdDel.Enabled = Not b
    cmdClear.Enabled = Not b
    cmdEdit.Enabled = Not b
    cmdAdd.Enabled = Not b
    txtCard.Enabled = Not b
    cmdAdd.Default = Not b
    cmdOK.Default = b
End Sub

Private Sub Form_Resize()
On Error Resume Next
    List1.Width = Width / 15 - 38
    List1.Height = Me.Height / 15 - 144
    PicTop.Width = Width / 15 - 16
    Cls
    Line (2, 2)-(Width / 15 - 14, Height / 15 - 29), 10921638, B
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
    With List1
        If (ColumnHeader.Index - 1) = .SortKey Then
            .SortOrder = 1 - .SortOrder
            .Sorted = True
        Else
            .Sorted = False
            .SortOrder = 0
            .SortKey = ColumnHeader.Index - 1
            .Sorted = True
        End If
    End With
End Sub

Private Sub List1_DblClick()
On Error GoTo aaaa
    Dim j As Long
    j = List1.SelectedItem.Index
    cmdEdit_Click
aaaa:
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo aaaa
    If KeyCode = vbKeyDelete Then
        Dim j As Long
        j = List1.SelectedItem.Index
        cmdDel_Click
    End If
aaaa:
End Sub

Private Sub txtCost_LostFocus()
On Error GoTo aaaa
    Dim c As Currency
    c = CCur(txtCost.Text)
Exit Sub
aaaa:
    txtCost.Text = txtCost.Tag
End Sub

Private Sub txtDate_LostFocus()
On Error GoTo aaaa
    Dim d As Date
    d = CDate(txtDate.Text)
Exit Sub
aaaa:
    txtDate.Text = txtDate.Tag
End Sub

