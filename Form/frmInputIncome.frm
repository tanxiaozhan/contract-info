VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputIncome 
   Caption         =   "数据修改"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   8670
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   9660
      Begin VB.Image imgIcon 
         Height          =   480
         Left            =   60
         Top             =   -15
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "数据修改"
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
         TabIndex        =   7
         Top             =   120
         Width           =   900
      End
   End
   Begin 合同管理.XPButton cmdExit 
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "返 回(&Q)"
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
   Begin 合同管理.XPButton cmdSave 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "保 存(&S)"
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
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   300
      Left            =   2595
      TabIndex        =   9
      Top             =   3030
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   25493505
      CurrentDate     =   39889
   End
   Begin 合同管理.XPButton cmdCLSDate 
      Height          =   330
      Left            =   2835
      TabIndex        =   10
      Top             =   3000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   582
      Caption         =   "清除"
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
   Begin 合同管理.FTextBox txtIncome 
      Height          =   300
      Index           =   1
      Left            =   1395
      TabIndex        =   0
      Top             =   3030
      Width           =   2055
      _ExtentX        =   3625
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
      Locked          =   -1  'True
   End
   Begin 合同管理.FTextBox txtIncome 
      Height          =   300
      Index           =   2
      Left            =   1395
      TabIndex        =   1
      Top             =   3750
      Width           =   2055
      _ExtentX        =   3625
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
   Begin 合同管理.FTextBox txtIncome 
      Height          =   300
      Index           =   3
      Left            =   1395
      TabIndex        =   2
      Top             =   4470
      Width           =   2055
      _ExtentX        =   3625
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
      isNumber        =   -1  'True
      MaxLength       =   15
   End
   Begin 合同管理.FTextBox txtIncome 
      Height          =   300
      Index           =   4
      Left            =   5835
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3030
      Width           =   2055
      _ExtentX        =   3625
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
   Begin 合同管理.FTextBox txtIncome 
      Height          =   300
      Index           =   5
      Left            =   5835
      TabIndex        =   4
      Top             =   3750
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   529
      BackColor       =   14737632
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
      Locked          =   -1  'True
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "保存成功!"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "借支情况"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label29 
      Caption         =   "录入日期"
      Height          =   160
      Left            =   4995
      TabIndex        =   23
      Top             =   3810
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "收款帐号"
      Height          =   160
      Left            =   5040
      TabIndex        =   22
      Top             =   3090
      Width           =   735
   End
   Begin VB.Label Label24 
      Caption         =   "收 款 人"
      Height          =   160
      Left            =   600
      TabIndex        =   21
      Top             =   3800
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "收款金额(元)"
      Height          =   165
      Left            =   240
      TabIndex        =   20
      Top             =   4515
      Width           =   1080
   End
   Begin VB.Label Label21 
      Caption         =   "收款日期"
      Height          =   160
      Left            =   600
      TabIndex        =   19
      Top             =   3075
      Width           =   735
   End
   Begin VB.Label Label30 
      Caption         =   "工程地点"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblHTMC"
      Height          =   160
      Index           =   2
      Left            =   1560
      TabIndex        =   17
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label32 
      Caption         =   "合同名称"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblGCDD"
      Height          =   160
      Index           =   3
      Left            =   5880
      TabIndex        =   15
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "合同编号"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblHTBH"
      Height          =   160
      Index           =   0
      Left            =   1560
      TabIndex        =   13
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label Label35 
      Caption         =   "委托单位"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblXMBH"
      Height          =   160
      Index           =   1
      Left            =   5880
      TabIndex        =   11
      Top             =   1680
      Width           =   2655
   End
End
Attribute VB_Name = "frmInputIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dblBorrow As Double          ' 借支金额


Private Sub cmdCLS_Click(Index As Integer)
    txtIncome(Index + 13).Text = ""
End Sub

Private Sub cmdCLSDate_Click()
    txtIncome(1).Text = ""
End Sub

Private Sub cobCBFS_Expand()
    cmdSave.Enabled = True
End Sub

Private Sub DTPicker1_CloseUp()
    txtIncome(14).Text = DTPicker1.value
End Sub

Private Sub DTPicker2_CloseUp()
    txtIncome(15).Text = DTPicker2.value
End Sub

Private Sub DTPicker3_CloseUp()
    txtIncome(16).Text = DTPicker3.value
End Sub

Private Sub DTPicker4_CloseUp()
    txtIncome(1).Text = DTPicker4.value
End Sub

Private Sub Form_Activate()
    If UCase(DataOperateState) = "EDIT" Then
         txtIncome(1).SetFocus
    End If

End Sub

Private Sub Form_Load()
'On Error GoTo aaaa
    'Me.WindowState = vbMaximized    '最大化窗口
    Width = 8790
    Height = 6795
    Me.Top = 500
    Me.Left = 1500
    imgIcon.Picture = frmMain.cmdLeft(2).Picture
    Label28.caption = ""
    txtIncome(5).Text = Now()   '录入日期
    
    
    Select Case DataOperateState
        Case "EDIT"    '编辑
            
            fillMain    '加载合同信息
            
            fillData
        
        Case "ADD"           '新增合同记录
            Me.caption = "数据录入"
            Label1.caption = Me.caption
            fillMain
            
    End Select
    
    
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical

End Sub
Sub fillData()      '借支信息
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim temp As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    strSQL = "select * from Income where id=" & incomeID
    rs.Open strSQL, Conn, 1, 1
    
    For i = 1 To 5
        If Not IsNull(rs.Fields(i).value) Then
            txtIncome(i).Text = rs.Fields(i).value
        End If
    Next
    rs.Close
    Conn.Close

End Sub
Sub fillMain()
    Dim strSQL As String
    Dim i As Integer
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    
    
    DBConnect
    strSQL = "select htbh,wtdw,htmc,gcdd from main where id=" & mainID
    rs.Open strSQL, Conn, 1, 1

    If Not rs.EOF Then
        For i = 0 To 3
            If IsNull(rs.Fields(i).value) Then
                lblHTinfo(i).caption = "未录入"
            Else
                lblHTinfo(i).caption = rs.Fields(i).value
            End If
        Next
    End If
    rs.Close
    Conn.Close
   

End Sub
Private Sub Form_Resize()
On Error Resume Next
    
    'PicTop.Width = Width * 0.5
    'PicTop.ScaleWidth = Me.ScaleWidth
    Cls
    'Line (2, 2)-(Width - 200, Height - 100), 10921638, B

End Sub

Private Sub Form_Unload(Cancel As Integer)
    DataOperateState = "ADD"
    frmList.Show   '显示列表
End Sub

Private Sub txtIncome_Change(Index As Integer)
    Label28.caption = ""
    cmdSave.Enabled = True
    
End Sub

Private Sub txtIncome_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Index < 4 Then
        txtIncome(Index + 1).SetFocus
    End If
End Sub
Private Sub cmdSave_Click()
   ' On Error GoTo errmsg
    
    If Trim(txtIncome(3).Text) = "" Then
        MsgBox "未填写[收款金额]", vbInformation, Me.caption
        txtIncome(3).SetFocus
        Exit Sub
    End If
    
    Select Case DataOperateState
        Case "EDIT"
            UpdateData "income", 5, incomeID
            
        Case "ADD"
            AddRecord "income", 5
            
            For i = 1 To 4
                txtIncome(i).Text = ""
            Next
            
            txtIncome(5).Text = Now()
            cmdSave.Enabled = False
            Label28.caption = "保存成功!"
            txtIncome(1).SetFocus
            
    End Select
        
    
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, Me.caption
    
End Sub
Sub UpdateData(strTable As String, byteFieldCount As Byte, lngID As Long)
    '参数:表名,字段数量,记录ID
    Dim strSQL As String
    Dim rs As ADODB.Recordset
                
    DBConnect
    
    Set rs = New ADODB.Recordset
    rs.Open "Select top 1 * From " & strTable, Conn, 1, 1
    For i = 1 To byteFieldCount
                    
        strText = Trim(txtIncome(i).Text)
        If strText = "" Then
            strText = "Null"
        Else
            If FieldTypeIsChar(rs.Fields(i).Type) Then     '字符或日期字段,加单引号
                strText = "'" & strText & "'"
            End If
        End If
                    
        strSQL = strSQL & rs.Fields(i).Name & "=" & strText & ","
    
    Next
    
    strSQL = Left(strSQL, Len(strSQL) - 1)
    strSQL = "update " & strTable & " set " & strSQL & " where id=" & lngID
    
    Conn.Execute strSQL
            
    Unload Me

End Sub
Sub AddRecord(strTable As String, byteFieldCount As Byte)    'flag -借支记录标志
    Dim strField, strValue, strText, strSQL As String
    Dim i, iCount As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    DBConnect
    strSQL = "select top 1  * from " & strTable
    rs.Open strSQL, Conn, 1, 1
    
    strField = "("
    strValue = "("
    For i = 1 To byteFieldCount
        strText = Trim(txtIncome(i).Text)
                                    
        If strText <> "" Then
            strField = strField & rs.Fields(i).Name & ","
            If FieldTypeIsChar(rs.Fields(i).Type) Then   '字符或日期字段,加单引号
                strText = "'" & strText & "'"
            End If
            strValue = strValue & strText & ","
         End If
    
    Next
    strField = strField & "zhtid)"
    strValue = strValue & mainID & ")"
    
    strSQL = "insert into " & strTable & " " & strField & "  values" & strValue
    'MsgBox strSQL
    Conn.Execute strSQL

End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

