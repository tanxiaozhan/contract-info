VERSION 5.00
Begin VB.Form frmDoc 
   Caption         =   "项目确认单"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   9150
   Begin 合同管理.XPButton cmdExit 
      Height          =   375
      Left            =   6960
      TabIndex        =   4
      Top             =   7800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "返  回(&Q)"
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
   Begin 合同管理.XPButton cmdOpen 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   7800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "在MS-WORD中打开(&O)"
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
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9660
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "生成项目确认单"
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
         TabIndex        =   1
         Top             =   120
         Width           =   1575
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   60
         Top             =   -15
         Width           =   480
      End
   End
   Begin VB.OLE oleWord 
      AutoActivate    =   0  'Manual
      BorderStyle     =   0  'None
      Class           =   "Word.Document.8"
      Height          =   7215
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   2
      Top             =   480
      Width           =   9135
   End
   Begin VB.Label lblInfo 
      Caption         =   "正在生成项目确认单数据，请稍候..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Width           =   5415
   End
End
Attribute VB_Name = "frmDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const OLE_CREATE_EMBED = 0
Const OLE_ACTIVATE = 7
Const wdReplaceAll = 2
Dim strDocName As String

Private Sub cmdExit_Click()
    frmMain.cmdLeft_Click 1
End Sub
Private Sub Form_Activate()
    On Error GoTo errmsg:
    Dim rs As ADODB.Recordset
    Dim strSQL, strField, strValue, strFormat As String
    Dim dblSum As Double '结算金额
    Dim dblBudget As Double '预算借支金额
        
        
    Select Case curDOCType
        Case 1:
            strDocName = "项目结算单"
            strSQL = "select main.wtdwlxr,main.wtdwlxdh,main.wtdw,main.htmc," & _
                     "sub.cbfs,sub.clr,sub.jcrs,sub.jcrq,sub.tcrq,subsec.gzny,sub.gcdd,subsec.htdj,subsec.sjgzl,sub.qt,sub.jsj" & " " & _
                     "from main,sub,subsec where sub.id=" & subID & " and main.id=sub.zhtid and sub.id=subsec.zhtid"
            strFormat = ",,,,,,,yyyy年m月d日,yyyy年m月d日,,,0.00,0.00,0.00,0.00"
        Case 2:
            strDocName = "项目确认单"
            strSQL = "select main.wtdwlxr,main.wtdwlxdh,main.wtdw,main.htmc," & _
                     "sub.cbfs,sub.clr,sub.jcrs,sub.jcrq,sub.tcrq,subsec.gzny,sub.gcdd,sub.jsj,sub.ysjzje,subsec.gzl*subsec.htdj as yssr" & " " & _
                     "from main,sub,subsec where sub.id=" & subID & " and main.id=sub.zhtid and sub.id=subsec.zhtid"
            strFormat = ",,,,,,,yyyy年m月d日,yyyy年m月d日,,,0.00,0.00,0.00"
    
        Case 3:
            strDocName = "项目借支单"
            strSQL = "select main.wtdwlxr,main.wtdwlxdh,main.wtdw,main.htmc," & _
                     "sub.cbfs,sub.clr,sub.jcrs,sub.jcrq,subsec.gzny,sub.gcdd" & " " & _
                     "from main,sub,subsec where sub.id=" & subID & " and main.id=sub.zhtid and sub.id=subsec.zhtid"
            strFormat = ",,,,,,,yyyy年m月d日,,"
    End Select
    
    Me.caption = strDocName
    Me.Label1.caption = "生成" & strDocName
    
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    rs.Open strSQL, Conn, 1, 1
        
    If rs.EOF Then
        MsgBox "数据库无指定的记录!", vbCritical, Me.caption
        
        frmMain.cmdLeft_Click 1  '返回列表窗口
        
        Exit Sub
    End If
    
    Me.lblInfo.caption = "正在生成" & strDocName & "，请稍侯..."
    
    wtdw = Trim(rs("wtdw"))
    clr = Trim(rs("clr"))
    
    oleWord.SourceDoc = GetApp & "Templets\" & strDocName & ".doc"
    oleWord.Action = OLE_CREATE_EMBED '创建嵌入（在OLE控制项中插入一个嵌入对象）
    
    oleWord.Action = 7     '激活OLE对象（用于编辑）
    oleWord.object.Application.Selection.Find.ClearFormatting
    oleWord.object.Application.Selection.Find.Replacement.ClearFormatting
    
    With oleWord.object.Application.Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    FormatString = Split(strFormat, ",")
    
    j = 0
        
    For i = 0 To rs.Fields.Count - 1
        
        If Not IsNull(rs.Fields(i).value) Then   '字段非空
            If UCase(rs.Fields(i).Name) = "CBFS" Then        '承包方式 0-再发包，1-自做
                strValue = IIf(rs.Fields(i).value = 0, "再发包", "自做")
            Else
                strValue = rs.Fields(i).value
            End If
        Else
            strValue = ""
        End If
        
        If FormatString(j) <> "" Then
            strValue = Format(strValue, FormatString(j))     '格式化数据
        End If
        
        With oleWord.object.Application.Selection.Find
            .Text = rs.Fields(i).Name
            .Replacement.Text = strValue
        End With
        oleWord.object.Application.Selection.Find.Execute Replace:=wdReplaceAll
        j = j + 1
    Next
    
    'dblSum = rs("jsj")          '结算价
    'dblBudget = rs("ysjzje")    '预算借支金额
    
    rs.Close
    
    Select Case curDOCType
        Case 1:      '结算单，增加一项人民币大写
        
        With oleWord.object.Application.Selection.Find
            .Text = "人民币大写"
            .Replacement.Text = coverToChinese(CStr(dblSum))
        End With
        oleWord.object.Application.Selection.Find.Execute Replace:=wdReplaceAll
    
    
        
        Case 2:   '项目确认单
            strSQL = "select ysjzje from sub where id=" & subID
            rs.Open strSQL, Conn, 1, 1
            If rs.EOF Then
                dblBudget = 0
            Else
                If Not IsNull(rs("ysjzje")) Then
                    dblBudget = rs("ysjzje")
                Else
                    dblBudget = 0
                End If
            End If
            
            rs.Close
            
            strSQL = "select jzrq,jzje,jzr,jzrzh,jzye from borrow where  zhtid=" & subID & " order by jzrq,lrrq"
            oleWord.object.Application.Selection.Find.Text = "Field100"
            oleWord.object.Application.Selection.Find.Execute
            rs.Open strSQL, Conn, 1, 1
            Do Until rs.EOF
                dblBudget = dblBudget - rs("jzje")
                If IsNull(rs("jzrq")) Then
                    temp = ""
                Else
                    temp = Format(CStr(rs("jzrq")), "yyyy年m月d日")
                End If
                oleWord.object.Application.Selection.TypeText Text:=temp
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                oleWord.object.Application.Selection.TypeText Text:=Format(CStr(rs("jzje")), "0.00")
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                If IsNull(rs("jzr")) Then
                    temp = ""
                Else
                    temp = CStr(rs("jzr"))
                End If
                
                oleWord.object.Application.Selection.TypeText Text:=temp
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                If IsNull(rs("jzrzh")) Then
                    temp = ""
                Else
                    temp = CStr(rs("jzrzh"))
                End If
                oleWord.object.Application.Selection.TypeText Text:=temp
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                oleWord.object.Application.Selection.TypeText Text:=Format(CStr(dblBudget), "0.00")
                rs.MoveNext
                If rs.EOF Then Exit Do
                oleWord.object.Application.Selection.InsertRowsBelow 1
            Loop
        Case 3:      '项目借支单
            strSQL = "select jzrq,jzje,jzr,jzrzh,jzye from borrow where id=" & borrowID
            oleWord.object.Application.Selection.Find.Text = "Field100"
            oleWord.object.Application.Selection.Find.Execute
            rs.Open strSQL, Conn, 1, 1
            Do Until rs.EOF
                If IsNull(rs("jzrq")) Then
                    temp = ""
                Else
                    temp = Format(CStr(rs("jzrq")), "yyyy年m月d日")
                End If
                oleWord.object.Application.Selection.TypeText Text:=temp
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                oleWord.object.Application.Selection.TypeText Text:=Format(CStr(rs("jzje")), "0.00")
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                If IsNull(rs("jzr")) Then
                    temp = ""
                Else
                    temp = CStr(rs("jzr"))
                End If
                
                oleWord.object.Application.Selection.TypeText Text:=temp
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                If IsNull(rs("jzrzh")) Then
                    temp = ""
                Else
                    temp = CStr(rs("jzrzh"))
                End If
                oleWord.object.Application.Selection.TypeText Text:=temp
                oleWord.object.Application.Selection.MoveRight Unit:=wdCharacter, Count:=1
                oleWord.object.Application.Selection.TypeText Text:=Format(CStr(dblBalace), "0.00")
                rs.MoveNext
                If rs.EOF Then Exit Do
                oleWord.object.Application.Selection.InsertRowsBelow 1
            Loop
    
    End Select
    
    cmdExit.SetFocus
    
    lblInfo.Visible = False
    oleWord.Visible = True
    strDocName = strDocName & "(" & wtdw & "--" & clr & ").doc"
    strDocName = Replace(strDocName, "\", ",")
    strDocName = Replace(strDocName, "/", ",")
    strDocName = Replace(strDocName, ":", ",")
    strDocName = Replace(strDocName, "*", ",")
    strDocName = Replace(strDocName, "?", ",")
    strDocName = Replace(strDocName, "<", ",")
    strDocName = Replace(strDocName, ">", ",")
    strDocName = Replace(strDocName, "|", ",")
    
    strDocName = GetApp & "DOC\" & strDocName
    
    If DirExists(GetApp & "Doc") = 0 Then     'DOC文件夹不存在则创建
        MkDir GetApp & "Doc"
    End If
    
    oleWord.object.Application.ActiveDocument.SaveAs strDocName
    oleWord.Action = 9
        
    cmdOpen.Enabled = True
    Exit Sub
errmsg:
    MsgBox "生成" & strDocName & "时发生错误。错误原因：" & Chr(13) & Err.Description & Chr(13) & "源文件：" & GetApp & "Templets\" & strDocName & ".doc", vbExclamation, "生成文档"
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo errmsg
    
    Dim wApp As Word.Application
    Set wApp = New Word.Application
    wApp.Documents.Open strDocName
    wApp.Visible = True
    Exit Sub
    
errmsg:
    
    MsgBox "打开word文档时发生错误。错误原因：" & Chr(13) & Err.Description, vbCritical, "错误"

End Sub

Private Sub Form_Load()
    imgIcon.Picture = frmMain.cmdLeft(CInt(curDOCType) + 2).Picture
    Me.Height = 8760
    Me.Width = 9270
    oleWord.Visible = False
    cmdOpen.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mainID = 0
End Sub
