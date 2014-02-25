VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportItem 
   Caption         =   "导出项目资料（新）"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   ControlBox      =   0   'False
   Icon            =   "frmExportItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11055
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   2805
      TabIndex        =   7
      Text            =   "正在导出数据，请稍候..."
      Top             =   2910
      Width           =   5880
   End
   Begin 合同管理.XPButton cmdCanel 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "取消选择"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin 合同管理.XPButton cmdAll 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "全 选"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin 合同管理.XPButton cmdExit 
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "返　回"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin 合同管理.XPButton cmdExport 
      Height          =   495
      Left            =   7815
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "导　出"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportItem.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportItem.frx":16D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportItem.frx":1C6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportItem.frx":2208
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExportItem.frx":27A2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView List1 
      Height          =   6105
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   10769
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "合同编号"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "委托单位"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "委托单位联系人"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "委托单位联系电话"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "合同名称"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "工程地点"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "测绘内容"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "合同总价(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "进场日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "退场日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "其它[补贴...](元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "结算价(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "结算日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "结余金额(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "付款方式"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "备注"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "录入日期"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   18
         Text            =   "项目负责人"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "请选择导出的项目："
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmExportItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAll_Click()
    For i = 1 To List1.ListItems.Count
        List1.ListItems(i).Selected = True
    Next
End Sub

Private Sub cmdCanel_Click()
    For i = 1 To List1.ListItems.Count
        List1.ListItems(i).Selected = False
    Next
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
'On Error GoTo errmsg
    Dim sql As String
    Dim id As String
    Dim rs As ADODB.Recordset
    Dim rsBorrow As ADODB.Recordset
    Dim jzelj As Double '借支额累计
    Dim isFill As Boolean
    Dim prevItemRow As Integer '上一合同起始行
    Dim itemChangeSum As Double '项目支出总费用
    Dim strgzny As String '工作内容
    Dim dblskhj As Double '收款合计
    Dim strskxx As String '收款信息
    Dim strskje As String '收款金额
    
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range

    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Selected Then Exit For
    Next
    
    If i > List1.ListItems.Count Then    '未选择的项目
        MsgBox "未选择项目，导出操作中止！", vbOKOnly, "导出项目资料"
        Exit Sub
    End If
    

    cmdAll.Visible = False
    cmdCanel.Visible = False
    cmdExport.Enabled = False
    
    Set rs = New ADODB.Recordset
    Set rsBorrow = New ADODB.Recordset
    DBConnect
    
    startRow = 3  '从第3行开始填充
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    Dlg.Filter = "MS Excel文件(*.xls)|*.xls"
    Dlg.FileName = "项目资料(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "导出项目资料"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    'strFormat = ";;;;;;;yyyy年mm月dd日;yyyy年mm月dd日;##,##0.00;yyyy年mm月dd日;##,##0.00;##,##0.00;##,##0.00;yyyy年mm月dd日"
    'arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\项目资料(新).xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    n = 0
    row = 1
    pbar.Max = 1
    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Selected Then pbar.Max = pbar.Max + 1
    Next
    If pbar.Max > 1 Then pbar.Max = pbar.Max - 1
    
    

    pbar.Visible = True '进度条
    Text1.Visible = True
    
    
        
    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Selected Then
            id = GetID(List1.ListItems(i).Key)
            sql = "select main.id as mainid,sub.id as subid ,main.wtdw,main.htmc,main.fzr,main.htzj as mainhtzj,main.jsrq as mainjsrq,main.jsj as mainjsj,sub.clr," & _
                        "sub.jcrq,sub.tcrq,sub.ysjzje,sub.jsj as subjsj,sub.jsrq as subjsrq " & _
                  "from main,sub " & _
                  "where main.id=" & id & " and main.id=sub.zhtid"
                  
            rs.Open sql, Conn, 1, 1
            isFill = False
            prevItemRow = row
            itemChangeSum = 0
            Do While Not rs.EOF
                If Not isFill Then
                    isFill = True
                    n = n + 1
                    xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '第4行，1列
                    xlSheet.Cells(startRow + row, 2) = rs("wtdw")
                    xlSheet.Cells(startRow + row, 3) = rs("htmc")
                    xlSheet.Cells(startRow + row, 4) = rs("fzr")
                    xlSheet.Cells(startRow + row, 5) = Format(rs("mainhtzj"), "##,##0.00")
                    xlSheet.Cells(startRow + row, 5).Select
                    With xlApp.Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = True     '缩小字体填充
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
    
                    
                    xlSheet.Cells(startRow + row, 6) = Format(rs("mainjsrq"), "yyyy年mm月dd日")
                    
                    'If rs("mainjsrq") <> "" Then
                    '    xlSheet.Cells(startRow + row, 6).Select
                    '    xlApp.Selection.Columns.AutoFit      最佳列宽
                    'End If
                    
                    xlSheet.Cells(startRow + row, 7) = Format(rs("mainjsj"), "##,##0.00")
                    If rs("mainjsj") <> "" Then xlSheet.Cells(startRow + row, 26) = Format(rs("mainjsj") * 0.2, "##,##0.00")
                    
                    mainjsj = IIf(rs("mainjsj") <> "", rs("mainjsj"), 0)
                    
                    '获取收款日期、金额
                    sql = "select skrq,skje from income where zhtid=" & rs("mainid")
                    rsBorrow.Open sql, Conn, 1, 1
                    dblskhj = 0 '收款合计
                    strskxx = "" '收款信息
                    strskje = "" '收款金额
                    Do While Not rsBorrow.EOF
                        If rsBorrow("skje") <> "" Then dblskhj = dblskhj + rsBorrow("skje")
                        strskxx = strskxx & Format(rsBorrow("skrq"), "yyyy年mm月dd日") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        strskje = strskje & Format(rsBorrow("skje"), "##,##0.00") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    
                        rsBorrow.MoveNext
                    Loop
                    rsBorrow.Close
                    If strskxx <> "" Then
                        strskxx = Left(strskxx, Len(strskxx) - 4)   '删除未尾的换行 chr(13) & chr(10) & chr(13) & chr(10)
                        xlSheet.Cells(startRow + row, 8) = strskxx
                        'xlSheet.Cells(startRow + row, 8).Select
                        'xlApp.Selection.Columns.AutoFit
                    End If
                    
                    If strskje <> "" Then
                        strskje = Left(strskje, Len(strskje) - 4)
                        xlSheet.Cells(startRow + row, 9) = strskje
                    End If
                
                    With xlSheet.Cells(startRow + row, 9)    '水平右对齐
                        .HorizontalAlignment = xlRight
                        .VerticalAlignment = xlCenter
                    End With
                
                    If dblskhj > 0 Then
                        xlSheet.Cells(startRow + row, 10) = Format(dblskhj, "##,##0.00")   '已收款
                        If rs("mainjsj") <> "" Then xlSheet.Cells(startRow + row, 11) = Format(rs("mainjsj") - dblskhj, "##,##0.00")  '未收款
                        
                        xlSheet.Cells(startRow + row, 10).Select
                        With xlApp.Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = True     '缩小字体填充
                            .ReadingOrder = xlContext
                            .MergeCells = False
                        End With
            
                    End If
                
                    
                    
                End If
                
                xlSheet.Cells(startRow + row, 13) = rs("clr")
                xlSheet.Cells(startRow + row, 14) = Format(rs("jcrq"), "yyyy年mm月dd日")
                'If rs("jcrq") <> "" Then
                '        xlSheet.Cells(startRow + row, 14).Select
                '        xlApp.Selection.Columns.AutoFit
                'End If
                xlSheet.Cells(startRow + row, 15) = Format(rs("tcrq"), "yyyy年mm月dd日")
                'If rs("tcrq") <> "" Then
                '        xlSheet.Cells(startRow + row, 15).Select
                '        xlApp.Selection.Columns.AutoFit
                'End If
                xlSheet.Cells(startRow + row, 16) = Format(rs("ysjzje"), "##,##0.00")
                xlSheet.Cells(startRow + row, 22) = Format(rs("subjsj"), "##,##0.00")
                xlSheet.Cells(startRow + row, 23) = Format(rs("subjsrq"), "yyyy年mm月dd日")
                'If rs("subjsrq") <> "" Then
                '        xlSheet.Cells(startRow + row, 23).Select
                '        xlApp.Selection.Columns.AutoFit
                'End If
                
                itemChangeSum = itemChangeSum + IIf(rs("subjsj") <> "", rs("subjsj"), 0)
                
                '获取工作内容
                sql = "select gzny from subsec where zhtID=" & rs("subid")
                rsBorrow.Open sql, Conn, 1, 1
                strgzny = ""
                No = 0
                Do While Not rsBorrow.EOF
                    No = No + 1
                    strgzny = strgzny & Trim(CStr(No)) & "." & rsBorrow("gzny") & Chr(13) & Chr(10)
                    rsBorrow.MoveNext
                Loop
                rsBorrow.Close
                If Len(strgzny) > 2 Then strgzny = Left(strgzny, Len(strgzny) - 2)  '删除未位回车chr(13)
                
                xlSheet.Cells(startRow + row, 12) = strgzny
                                
                With xlSheet.Cells(startRow + row, 12)    '水平左对齐
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                End With
                
                
                '获取借支信息
                sql = "select * from borrow where zhtid=" & rs("subid") & " order by jzrq"
                rsBorrow.Open sql, Conn, 1, 1
                jzelj = 0 '借支额累计
                Do While Not rsBorrow.EOF
                    xlSheet.Cells(startRow + row, 17) = Format(rsBorrow("jzrq"), "yyyy年mm月dd日")
                    'If rsBorrow("jzrq") <> "" Then
                    '    xlSheet.Cells(startRow + row, 17).Select
                    '    xlApp.Selection.Columns.AutoFit
                    'End If
                    xlSheet.Cells(startRow + row, 18) = rsBorrow("jzr")
                    xlSheet.Cells(startRow + row, 19) = Format(rsBorrow("jzje"), "##,##0.00")
                    xlSheet.Cells(startRow + row, 20) = Format(rsBorrow("jzye"), "##,##0.00")
                    jzelj = jzelj + rsBorrow("jzje")
                                        
                    rsBorrow.MoveNext
                    row = row + 1
                Loop
                
                If rsBorrow.RecordCount < 1 Then row = row + 1
                
                If jzelj > 0 Then xlSheet.Cells(startRow + row - 1, 21) = Format(jzelj, "##,##0.00")    '借支额累计
                
                If rsBorrow.RecordCount > 0 Then   '有借支记录则合并单元格
                    For j = 12 To 16
                        xlSheet.Range(xlSheet.Cells(startRow + row - rsBorrow.RecordCount, j), xlSheet.Cells(startRow + row - 1, j)).Merge
                    Next
                    For j = 21 To 23
                        xlSheet.Range(xlSheet.Cells(startRow + row - rsBorrow.RecordCount, j), xlSheet.Cells(startRow + row - 1, j)).Merge
                    Next
                End If
                
                rsBorrow.Close
                
                rs.MoveNext
            Loop
            
            
            recount = rs.RecordCount
            rs.Close
            
            If recount < 1 Then      '没有子合同的情况,直接从主合同表中取得数据。
                sql = "select id,wtdw,htmc,fzr,htzj,jsrq,jcrq,tcrq,jsj " & _
                  "from main " & _
                  "where main.id=" & id
                  
                rs.Open sql, Conn, 1, 1
                Do While Not rs.EOF
                    
                    n = n + 1
                    
                    xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '第4行，1列
                    xlSheet.Cells(startRow + row, 2) = rs("wtdw")
                    xlSheet.Cells(startRow + row, 3) = rs("htmc")
                    xlSheet.Cells(startRow + row, 5) = Format(rs("htzj"), "##,##0.00")
                    xlSheet.Cells(startRow + row, 5).Select
                    With xlApp.Selection
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = True     '缩小字体填充
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
    
                    
                    xlSheet.Cells(startRow + row, 6) = Format(rs("jsrq"), "yyyy年mm月dd日")
                    If (rs("jsj")) <> "" Then
                        xlSheet.Cells(startRow + row, 7) = Format(rs("jsj"), "##,##0.00")
                        xlSheet.Cells(startRow + row, 26) = Format(rs("jsj") * 0.2, "##,##0.00")
                        xlSheet.Cells(startRow + row, 25) = Format(rs("jsj"), "##,##0.00")
                        xlSheet.Cells(startRow + row, 27).FormulaR1C1 = "=RC[-2]-RC[-1]"
                        
                    End If
                    xlSheet.Cells(startRow + row, 14) = Format(rs("jcrq"), "yyyy年mm月dd日")
                    xlSheet.Cells(startRow + row, 15) = Format(rs("tcrq"), "yyyy年mm月dd日")
                    
                    '获取收款日期、金额
                    sql = "select skrq,skje from income where zhtid=" & rs("id")
                    rsBorrow.Open sql, Conn, 1, 1
                    dblskhj = 0 '收款合计
                    strskxx = "" '收款信息
                    strskje = "" '收款金额
                    Do While Not rsBorrow.EOF
                        If rsBorrow("skje") <> "" Then dblskhj = dblskhj + rsBorrow("skje")
                        strskxx = strskxx & Format(rsBorrow("skrq"), "yyyy年mm月dd日") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        strskje = strskje & Format(rsBorrow("skje"), "##,##0.00") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    
                        rsBorrow.MoveNext
                    Loop
                    rsBorrow.Close
                    If strskxx <> "" Then
                        strskxx = Left(strskxx, Len(strskxx) - 4)   '删除未尾的换行 chr(13) & chr(10) & chr(13) & chr(10)
                        xlSheet.Cells(startRow + row, 8) = strskxx
                    End If
                    
                    If strskje <> "" Then
                        strskje = Left(strskje, Len(strskje) - 4)
                        xlSheet.Cells(startRow + row, 9) = strskje
                    End If
                
                    With xlSheet.Cells(startRow + row, 9)    '水平右对齐
                        .HorizontalAlignment = xlRight
                        .VerticalAlignment = xlCenter
                    End With
                
                    If dblskhj > 0 Then
                        xlSheet.Cells(startRow + row, 10) = Format(dblskhj, "##,##0.00")   '已收款
                        If rs("jsj") <> "" Then xlSheet.Cells(startRow + row, 11) = Format(rs("jsj") - dblskhj, "##,##0.00")  '未收款
                        
                        xlSheet.Cells(startRow + row, 10).Select
                        With xlApp.Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = True     '缩小字体填充
                            .ReadingOrder = xlContext
                            .MergeCells = False
                        End With
            
                    End If
                    
                    rs.MoveNext
                Loop
                rs.Close
                row = row + 1
            Else
            
            
                If itemChangeSum > 0 Then xlSheet.Cells(startRow + row - 1, 24) = Format(itemChangeSum, "##,##0.00")   '本项目支出费用合计
                itemChangeSum = mainjsj - itemChangeSum   '本项目结算后剩余
                If itemChangeSum > 0 Then xlSheet.Cells(startRow + row - 1, 25) = Format(itemChangeSum, "##,##0.00")
            
                For j = 1 To 11
                   xlSheet.Range(xlSheet.Cells(startRow + prevItemRow, j), xlSheet.Cells(startRow + row - 1, j)).Merge
                Next
                For j = 24 To 28
                    xlSheet.Range(xlSheet.Cells(startRow + prevItemRow, j), xlSheet.Cells(startRow + row - 1, j)).Merge
                Next
            
                xlSheet.Cells(startRow + prevItemRow, 27).FormulaR1C1 = "=RC[-2]-RC[-1]"
            
            End If
            

        End If
        
        If pbar.value < pbar.Max Then pbar.value = pbar.value + 1
        
    Next
    

    Set xlRange = xlSheet.Range(xlSheet.Cells(startRow, 1), xlSheet.Cells(startRow + row - 1, 28))
    
    With xlRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    xlRange.rowHeight = 25
    Set xlRange = xlSheet.Range(xlSheet.Cells(1, 4), xlSheet.Cells(1, 4))
    xlRange.ColumnWidth = 0
    Set xlRange = xlSheet.Range(xlSheet.Cells(1, 20), xlSheet.Cells(1, 20))
    xlRange.ColumnWidth = 0
    
    
    xlSheet.Cells(4, 1).Select
    
    
    xlApp.DisplayAlerts = False   '保存不显示覆盖提示
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    pbar.value = pbar.Max
    
    MsgBox "项目资料导出完成！" & Chr(13) & "保存到" & Dlg.FileName, vbInformation, "导出项目资料"
    
    pbar.Visible = False
    Text1.Visible = False
    cmdAll.Visible = True
    cmdCanel.Visible = True
    cmdExport.Enabled = True
    
    GoTo end_sub

errmsg:
    pbar.Visible = False
    cmdAll.Visible = True
    cmdCanel.Visible = True
    If Not IsNull(xlApp) Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    
    If Err.Number = 32755 Then Exit Sub     '32755，用户点击取消按钮
    MsgBox Err.Description, vbInformation, "导出项目资料"

end_sub:
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Visible = False
    pbar.Visible = False
    loadList
    List1.BackColor = RGB(225, 225, 225)
    
End Sub
Private Sub loadList()
    Dim No, htlx As Integer
    Dim Item As ListItem
    Dim rs As ADODB.Recordset
    Dim strFormat As String
    
    'strFormat = "0;1;2;3;4;5;6;7;8;##,##0.00;yyyy-mm-dd;yyyy-mm-dd;##,##0.00;##,##0.00;yyyy-mm-dd;##,##0.00;16;yyyy-mm-dd hh:mm:ss"
    strFormat = ";;;;;;;;;##,##0.00;yyyy-mm-dd;yyyy-mm-dd;##,##0.00;##,##0.00;yyyy-mm-dd;##,##0.00;;;yyyy-mm-dd hh:mm:ss"
    strFormat = Replace(strFormat, ";0.000", GetDecLen(bytAfterDec))
    arrayFormat = Split(strFormat, ";")
    Set rs = New ADODB.Recordset
    DBConnect
    rs.Open "Select * From main order by lrrq desc,htbh", Conn, 1, 1
    List1.ListItems.Clear
    No = 0
    
    Do While Not rs.EOF
        No = No + 1
        
        
        If rs("htlx") > 2 Then     '合同类型
            htlx = 2
        Else
            htlx = rs("htlx")
        End If
        
        Set Item = List1.ListItems.Add(, Trim(CStr(rs("id"))) & "k", No, , htlx + 1)
            
        For i = 2 To rs.Fields.Count - 3
            
            If IsNull(rs.Fields(i).value) Then
                temp = " "
            Else
                temp = rs.Fields(i).value
                If Not FieldTypeIsChar(rs.Fields(i).Type) Then
                    If temp = 0 Then
                        temp = " "
                    End If
                End If
            End If
            
            Item.SubItems(i - 1) = Format(temp, arrayFormat(i))
            
        Next
        
        If Not IsNull(rs("fzr").value) Then Item.SubItems(18) = rs("fzr").value
        
        
        
        If rs("yjs") Then
            textcolor = color(2)
        Else
            textcolor = RGB(0, 0, 0)
        End If
        
        For i = 1 To List1.ColumnHeaders.Count - 2
            List1.ListItems(No).ForeColor = textcolor
            List1.ListItems(No).ListSubItems.Item(i).ForeColor = textcolor
        Next
        
        rs.MoveNext
        
    Loop
    
    cmdExport.Enabled = IIf(rs.RecordCount > 0, True, False)
    rs.Close
    Set rs = Nothing
    Conn.Close
    Set Conn = Nothing
End Sub
Function GetDecLen(AfterDec As Byte)
    Dim i As Byte
    Dim strDecFormat As String
        
    strDecFormat = "0."
    
    If AfterDec = 0 Then
        strDecFormat = "0"
    Else
        For i = 1 To AfterDec
            strDecFormat = strDecFormat & "0"
        Next
    End If
    
    GetDecLen = ";" & strDecFormat
    
End Function
