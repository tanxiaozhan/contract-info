VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportItem 
   Caption         =   "������Ŀ���ϣ��£�"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   ControlBox      =   0   'False
   Icon            =   "frmExportItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11055
   StartUpPosition =   1  '����������
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
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
      Text            =   "���ڵ������ݣ����Ժ�..."
      Top             =   2910
      Width           =   5880
   End
   Begin ��ͬ����.XPButton cmdCanel 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "ȡ��ѡ��"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin ��ͬ����.XPButton cmdAll 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "ȫ ѡ"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin ��ͬ����.XPButton cmdExit 
      Height          =   495
      Left            =   9480
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "������"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin ��ͬ����.XPButton cmdExport 
      Height          =   495
      Left            =   7815
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "������"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Text            =   "���"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "��ͬ���"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "ί�е�λ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "ί�е�λ��ϵ��"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "ί�е�λ��ϵ�绰"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "��ͬ����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "���̵ص�"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "�������"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "��ͬ�ܼ�(Ԫ)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "��������"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "�˳�����"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "����[����...](Ԫ)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "�����(Ԫ)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "��������"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "������(Ԫ)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "���ʽ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "��ע"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Text            =   "¼������"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   18
         Text            =   "��Ŀ������"
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
      Caption         =   "��ѡ�񵼳�����Ŀ��"
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
    Dim jzelj As Double '��֧���ۼ�
    Dim isFill As Boolean
    Dim prevItemRow As Integer '��һ��ͬ��ʼ��
    Dim itemChangeSum As Double '��Ŀ֧���ܷ���
    Dim strgzny As String '��������
    Dim dblskhj As Double '�տ�ϼ�
    Dim strskxx As String '�տ���Ϣ
    Dim strskje As String '�տ���
    
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range

    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Selected Then Exit For
    Next
    
    If i > List1.ListItems.Count Then    'δѡ�����Ŀ
        MsgBox "δѡ����Ŀ������������ֹ��", vbOKOnly, "������Ŀ����"
        Exit Sub
    End If
    

    cmdAll.Visible = False
    cmdCanel.Visible = False
    cmdExport.Enabled = False
    
    Set rs = New ADODB.Recordset
    Set rsBorrow = New ADODB.Recordset
    DBConnect
    
    startRow = 3  '�ӵ�3�п�ʼ���
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    Dlg.Filter = "MS Excel�ļ�(*.xls)|*.xls"
    Dlg.FileName = "��Ŀ����(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "������Ŀ����"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    'strFormat = ";;;;;;;yyyy��mm��dd��;yyyy��mm��dd��;##,##0.00;yyyy��mm��dd��;##,##0.00;##,##0.00;##,##0.00;yyyy��mm��dd��"
    'arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\��Ŀ����(��).xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    n = 0
    row = 1
    pbar.Max = 1
    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Selected Then pbar.Max = pbar.Max + 1
    Next
    If pbar.Max > 1 Then pbar.Max = pbar.Max - 1
    
    

    pbar.Visible = True '������
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
                    xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '��4�У�1��
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
                        .ShrinkToFit = True     '��С�������
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
    
                    
                    xlSheet.Cells(startRow + row, 6) = Format(rs("mainjsrq"), "yyyy��mm��dd��")
                    
                    'If rs("mainjsrq") <> "" Then
                    '    xlSheet.Cells(startRow + row, 6).Select
                    '    xlApp.Selection.Columns.AutoFit      ����п�
                    'End If
                    
                    xlSheet.Cells(startRow + row, 7) = Format(rs("mainjsj"), "##,##0.00")
                    If rs("mainjsj") <> "" Then xlSheet.Cells(startRow + row, 26) = Format(rs("mainjsj") * 0.2, "##,##0.00")
                    
                    mainjsj = IIf(rs("mainjsj") <> "", rs("mainjsj"), 0)
                    
                    '��ȡ�տ����ڡ����
                    sql = "select skrq,skje from income where zhtid=" & rs("mainid")
                    rsBorrow.Open sql, Conn, 1, 1
                    dblskhj = 0 '�տ�ϼ�
                    strskxx = "" '�տ���Ϣ
                    strskje = "" '�տ���
                    Do While Not rsBorrow.EOF
                        If rsBorrow("skje") <> "" Then dblskhj = dblskhj + rsBorrow("skje")
                        strskxx = strskxx & Format(rsBorrow("skrq"), "yyyy��mm��dd��") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        strskje = strskje & Format(rsBorrow("skje"), "##,##0.00") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    
                        rsBorrow.MoveNext
                    Loop
                    rsBorrow.Close
                    If strskxx <> "" Then
                        strskxx = Left(strskxx, Len(strskxx) - 4)   'ɾ��δβ�Ļ��� chr(13) & chr(10) & chr(13) & chr(10)
                        xlSheet.Cells(startRow + row, 8) = strskxx
                        'xlSheet.Cells(startRow + row, 8).Select
                        'xlApp.Selection.Columns.AutoFit
                    End If
                    
                    If strskje <> "" Then
                        strskje = Left(strskje, Len(strskje) - 4)
                        xlSheet.Cells(startRow + row, 9) = strskje
                    End If
                
                    With xlSheet.Cells(startRow + row, 9)    'ˮƽ�Ҷ���
                        .HorizontalAlignment = xlRight
                        .VerticalAlignment = xlCenter
                    End With
                
                    If dblskhj > 0 Then
                        xlSheet.Cells(startRow + row, 10) = Format(dblskhj, "##,##0.00")   '���տ�
                        If rs("mainjsj") <> "" Then xlSheet.Cells(startRow + row, 11) = Format(rs("mainjsj") - dblskhj, "##,##0.00")  'δ�տ�
                        
                        xlSheet.Cells(startRow + row, 10).Select
                        With xlApp.Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = True     '��С�������
                            .ReadingOrder = xlContext
                            .MergeCells = False
                        End With
            
                    End If
                
                    
                    
                End If
                
                xlSheet.Cells(startRow + row, 13) = rs("clr")
                xlSheet.Cells(startRow + row, 14) = Format(rs("jcrq"), "yyyy��mm��dd��")
                'If rs("jcrq") <> "" Then
                '        xlSheet.Cells(startRow + row, 14).Select
                '        xlApp.Selection.Columns.AutoFit
                'End If
                xlSheet.Cells(startRow + row, 15) = Format(rs("tcrq"), "yyyy��mm��dd��")
                'If rs("tcrq") <> "" Then
                '        xlSheet.Cells(startRow + row, 15).Select
                '        xlApp.Selection.Columns.AutoFit
                'End If
                xlSheet.Cells(startRow + row, 16) = Format(rs("ysjzje"), "##,##0.00")
                xlSheet.Cells(startRow + row, 22) = Format(rs("subjsj"), "##,##0.00")
                xlSheet.Cells(startRow + row, 23) = Format(rs("subjsrq"), "yyyy��mm��dd��")
                'If rs("subjsrq") <> "" Then
                '        xlSheet.Cells(startRow + row, 23).Select
                '        xlApp.Selection.Columns.AutoFit
                'End If
                
                itemChangeSum = itemChangeSum + IIf(rs("subjsj") <> "", rs("subjsj"), 0)
                
                '��ȡ��������
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
                If Len(strgzny) > 2 Then strgzny = Left(strgzny, Len(strgzny) - 2)  'ɾ��δλ�س�chr(13)
                
                xlSheet.Cells(startRow + row, 12) = strgzny
                                
                With xlSheet.Cells(startRow + row, 12)    'ˮƽ�����
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                End With
                
                
                '��ȡ��֧��Ϣ
                sql = "select * from borrow where zhtid=" & rs("subid") & " order by jzrq"
                rsBorrow.Open sql, Conn, 1, 1
                jzelj = 0 '��֧���ۼ�
                Do While Not rsBorrow.EOF
                    xlSheet.Cells(startRow + row, 17) = Format(rsBorrow("jzrq"), "yyyy��mm��dd��")
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
                
                If jzelj > 0 Then xlSheet.Cells(startRow + row - 1, 21) = Format(jzelj, "##,##0.00")    '��֧���ۼ�
                
                If rsBorrow.RecordCount > 0 Then   '�н�֧��¼��ϲ���Ԫ��
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
            
            If recount < 1 Then      'û���Ӻ�ͬ�����,ֱ�Ӵ�����ͬ����ȡ�����ݡ�
                sql = "select id,wtdw,htmc,fzr,htzj,jsrq,jcrq,tcrq,jsj " & _
                  "from main " & _
                  "where main.id=" & id
                  
                rs.Open sql, Conn, 1, 1
                Do While Not rs.EOF
                    
                    n = n + 1
                    
                    xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '��4�У�1��
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
                        .ShrinkToFit = True     '��С�������
                        .ReadingOrder = xlContext
                        .MergeCells = False
                    End With
    
                    
                    xlSheet.Cells(startRow + row, 6) = Format(rs("jsrq"), "yyyy��mm��dd��")
                    If (rs("jsj")) <> "" Then
                        xlSheet.Cells(startRow + row, 7) = Format(rs("jsj"), "##,##0.00")
                        xlSheet.Cells(startRow + row, 26) = Format(rs("jsj") * 0.2, "##,##0.00")
                        xlSheet.Cells(startRow + row, 25) = Format(rs("jsj"), "##,##0.00")
                        xlSheet.Cells(startRow + row, 27).FormulaR1C1 = "=RC[-2]-RC[-1]"
                        
                    End If
                    xlSheet.Cells(startRow + row, 14) = Format(rs("jcrq"), "yyyy��mm��dd��")
                    xlSheet.Cells(startRow + row, 15) = Format(rs("tcrq"), "yyyy��mm��dd��")
                    
                    '��ȡ�տ����ڡ����
                    sql = "select skrq,skje from income where zhtid=" & rs("id")
                    rsBorrow.Open sql, Conn, 1, 1
                    dblskhj = 0 '�տ�ϼ�
                    strskxx = "" '�տ���Ϣ
                    strskje = "" '�տ���
                    Do While Not rsBorrow.EOF
                        If rsBorrow("skje") <> "" Then dblskhj = dblskhj + rsBorrow("skje")
                        strskxx = strskxx & Format(rsBorrow("skrq"), "yyyy��mm��dd��") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                        strskje = strskje & Format(rsBorrow("skje"), "##,##0.00") & Chr(13) & Chr(10) & Chr(13) & Chr(10)
                    
                        rsBorrow.MoveNext
                    Loop
                    rsBorrow.Close
                    If strskxx <> "" Then
                        strskxx = Left(strskxx, Len(strskxx) - 4)   'ɾ��δβ�Ļ��� chr(13) & chr(10) & chr(13) & chr(10)
                        xlSheet.Cells(startRow + row, 8) = strskxx
                    End If
                    
                    If strskje <> "" Then
                        strskje = Left(strskje, Len(strskje) - 4)
                        xlSheet.Cells(startRow + row, 9) = strskje
                    End If
                
                    With xlSheet.Cells(startRow + row, 9)    'ˮƽ�Ҷ���
                        .HorizontalAlignment = xlRight
                        .VerticalAlignment = xlCenter
                    End With
                
                    If dblskhj > 0 Then
                        xlSheet.Cells(startRow + row, 10) = Format(dblskhj, "##,##0.00")   '���տ�
                        If rs("jsj") <> "" Then xlSheet.Cells(startRow + row, 11) = Format(rs("jsj") - dblskhj, "##,##0.00")  'δ�տ�
                        
                        xlSheet.Cells(startRow + row, 10).Select
                        With xlApp.Selection
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .WrapText = False
                            .Orientation = 0
                            .AddIndent = False
                            .IndentLevel = 0
                            .ShrinkToFit = True     '��С�������
                            .ReadingOrder = xlContext
                            .MergeCells = False
                        End With
            
                    End If
                    
                    rs.MoveNext
                Loop
                rs.Close
                row = row + 1
            Else
            
            
                If itemChangeSum > 0 Then xlSheet.Cells(startRow + row - 1, 24) = Format(itemChangeSum, "##,##0.00")   '����Ŀ֧�����úϼ�
                itemChangeSum = mainjsj - itemChangeSum   '����Ŀ�����ʣ��
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
    
    
    xlApp.DisplayAlerts = False   '���治��ʾ������ʾ
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    pbar.value = pbar.Max
    
    MsgBox "��Ŀ���ϵ�����ɣ�" & Chr(13) & "���浽" & Dlg.FileName, vbInformation, "������Ŀ����"
    
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
    
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "������Ŀ����"

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
        
        
        If rs("htlx") > 2 Then     '��ͬ����
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
