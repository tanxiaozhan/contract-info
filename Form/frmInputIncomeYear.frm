VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExportIncomeYear 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "导出收款一览表"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4515
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.DTPicker DTPickerEnd 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   66846721
      CurrentDate     =   41682
   End
   Begin MSComCtl2.DTPicker DTPickerBegin 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   66846721
      CurrentDate     =   41682
   End
   Begin VB.CheckBox chkYear 
      Caption         =   "指定起止日期"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   420
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   4020
      Top             =   2085
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin 合同管理.FTextBox txtyear 
      Height          =   300
      Left            =   2460
      TabIndex        =   2
      Top             =   765
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      BackColor       =   12636398
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
      isNumber        =   -1  'True
      MaxLength       =   5
      afterdecimal    =   0
   End
   Begin 合同管理.XPButton cmdExport 
      Height          =   450
      Left            =   570
      TabIndex        =   1
      Top             =   1890
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   794
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
   Begin 合同管理.XPButton XPButton1 
      Height          =   450
      Left            =   2520
      TabIndex        =   0
      Top             =   1890
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   794
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
   Begin VB.Label lblEnd 
      Caption         =   "结束日期:"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblBegin 
      Caption         =   "起始日期:"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblYear 
      Caption         =   "年"
      Height          =   240
      Left            =   3270
      TabIndex        =   7
      Top             =   820
      Width           =   420
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "正在导出收款情况一览表，请稍候..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   450
      Width           =   4335
   End
   Begin VB.Label lblTip 
      Caption         =   "提示：不输入年份可导出所有年份数据"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   705
      TabIndex        =   4
      Top             =   1455
      Width           =   3075
   End
   Begin VB.Label lblInput 
      Caption         =   "请输入年份(4位):"
      Height          =   270
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1665
   End
End
Attribute VB_Name = "frmExportIncomeYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkYear_Click()
    If chkYear.value = 0 Then
        lblBegin.Visible = False
        lblEnd.Visible = False
        DTPickerBegin.Visible = False
        DTPickerEnd.Visible = False
        
        lblInput.Visible = True
        txtyear.Visible = True
        lblTip.Visible = True
        lblYear.Visible = True
    Else
        lblBegin.Visible = True
        lblEnd.Visible = True
        DTPickerBegin.Visible = True
        DTPickerEnd.Visible = True
        
        lblInput.Visible = False
        txtyear.Visible = False
        lblTip.Visible = False
        lblYear.Visible = False

    End If
End Sub

Private Sub Form_Activate()
    txtyear.SetFocus
    lblInfo.Visible = False
    pbar.Visible = False
    
    lblBegin.Visible = False
    lblEnd.Visible = False
    DTPickerBegin.Visible = False
    DTPickerEnd.Visible = False
    
    DTPickerBegin.value = CDate(str(Year(Date)) & "-1-1")

End Sub

Private Sub XPButton1_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
        
    On Error GoTo errmsg
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range
    Dim rs As ADODB.Recordset
    Dim rsIncome As ADODB.Recordset
    Dim strSQL As String
    Dim i, row, startRow, n As Integer
    Dim strFormat As String
    Dim strHTBH, strXMBH As String '合同编号,项目编号
    Dim dblTotal As Double    '借支余额
    
    startRow = 3  '从第3行开始填充
    
    Set rs = New ADODB.Recordset
    Set rsIncome = New ADODB.Recordset
    DBConnect
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    strSQL = "select  id,htbh,htmc,jcrq,tcrq,htzj,jsj" & " " & _
             "from main" & " " & _
             "order by main.lrrq desc"
    
    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        MsgBox "未找到相关记录，导出中止！", vbExclamation, "导出收款情况一览表"
        rs.Close
        Conn.Close
        Exit Sub
    End If
    
    lblBegin.Visible = False
    lblEnd.Visible = False
    DTPickerBegin.Visible = False
    DTPickerEnd.Visible = False
        
    lblInput.Visible = False
    txtyear.Visible = False
    lblTip.Visible = False
    lblYear.Visible = False

    
    lblInfo.Visible = True
    pbar.Visible = True
    pbar.Max = rs.RecordCount
    
    skYear = ""
   
    If chkYear.value = 1 Then   '指定起止日期
        skYear = DTPickerBegin.value & "至" & DTPickerEnd.value
    
    Else                        '只指定年份或全部
    
        If Trim(txtyear.Text) <> "" Then
            skYear = Trim(txtyear.Text) & "年"
        End If
    
    End If
    
    Dlg.Filter = "MS Excel文件(*.xls)|*.xls"
    Dlg.FileName = skYear & "收款情况一览表(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = skYear & "导出收款情况一览表"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    strFormat = ";;;yyyy年mm月dd日;yyyy年mm月dd日;##,##0.00;##,##0.00;yyyy年mm月dd日;##,##0.00;##,##0.00"
    arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\收款情况一览表.xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    strXMBH = ""    '项目编号
    strHTBH = ""   '合同编号
    n = 0
    row = 0
    
    
    
    Do While Not rs.EOF
        n = n + 1
        
        
        strSQL = "select skrq,skje from income where zhtid=" & rs("id")
        
        If chkYear.value = 1 Then    '指定起止日期
            strSQL = strSQL & " " & "and skrq>=#" & DTPickerBegin.value & "# and skrq<=#" & DTPickerEnd.value & "#"
        Else
        
            If Trim(txtyear.Text) <> "" Then
                strSQL = strSQL & " " & "  and skrq like '" & Trim(txtyear.Text) & "%'"
            End If
        
        End If
    
        strSQL = strSQL & " " & "order by skrq"
        
        rsIncome.Open strSQL, Conn, 1, 1
            
    
        If rsIncome.RecordCount > 0 Then
            
            xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '第4行，1列
        
            If IsNull(rs("jsj")) Then      '结算价
                dblTotal = 0
            Else
                dblTotal = CDbl(rs("jsj"))
            End If
            
            For i = 1 To 6 '1-合同编号,....
                If Not IsNull(rs.Fields(i).value) Then
                    xlSheet.Cells(startRow + row, 1 + i) = IIf(arrayFormat(i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(i)), rs.Fields(i).value)
                End If
            Next
            
            
            Do While Not rsIncome.EOF
            
                For i = 0 To 1    '收款情况
                    If Not IsNull(rsIncome.Fields(i).value) Then
                        xlSheet.Cells(startRow + row, 8 + i) = IIf(arrayFormat(7 + i) <> "", Format(CStr(rsIncome.Fields(i).value), arrayFormat(7 + i)), rsIncome.Fields(i).value)
                    End If
                
                Next
            
                If Not IsNull(rsIncome("skje")) Then    '计算收款余额
                    dblTotal = dblTotal - CDbl(rsIncome("skje"))
                End If
                If dblTotal < 0 Then
                    xlSheet.Cells(startRow + row, 10) = "未结算"
                Else
                    xlSheet.Cells(startRow + row, 10) = IIf(arrayFormat(9) <> "", Format(CStr(dblTotal), arrayFormat(9)), CStr(dblTotal))
                End If
                rsIncome.MoveNext
                row = row + 1
            Loop
            
            If rsIncome.RecordCount > 1 Then
                For i = 1 To 7
                    xlSheet.Range(xlSheet.Cells(startRow + row - 1, i), xlSheet.Cells(startRow + row - rsIncome.RecordCount, i)).Merge
                Next
                xlSheet.Range(xlSheet.Cells(startRow + row - 1, 11), xlSheet.Cells(startRow + row - rsIncome.RecordCount, 11)).Merge
            
            End If
        
        End If
        
        rsIncome.Close
        
        rs.MoveNext
        pbar.value = pbar.value + 1  '更新进度条
    Loop
    
    lblInfo.caption = "正在整理导出的数据格式..."
    
    pbar.value = 0
    pbar.Max = 6
    
    DoEvents
    
    Set xlRange = xlSheet.Range(xlSheet.Cells(startRow, 1), xlSheet.Cells(startRow + row - 1, 11))
    
    With xlRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pbar.value = pbar.value + 1
    
    With xlRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pbar.value = pbar.value + 1
    
    With xlRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pbar.value = pbar.value + 1
    
    With xlRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pbar.value = pbar.value + 1
    
    With xlRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pbar.value = pbar.value + 1
    
    With xlRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    pbar.value = pbar.value + 1
    
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    lblInfo.caption = "收款情况一览表导出成功！"
    lblInfo.Refresh
    
    MsgBox "收款情况一览表导出完成！" & Chr(13) & "保存到" & Dlg.FileName, vbInformation, "导出收款情况一览表"
    
    GoTo end_sub

errmsg:
    If Not (Err.Number = 32755 Or Err.Number = 1004) Then     '32755，用户点击取消按钮, 1004另存时选“否”或“取消”按钮
        MsgBox Err.Description, vbInformation, "导出收款情况一览表"
    End If

end_sub:
    Unload Me
    
End Sub
