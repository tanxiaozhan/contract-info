VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputBorrow 
   Caption         =   "数据修改"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8580
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   10
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
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
   End
   Begin 合同管理.XPButton cmdExit 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   5280
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
      TabIndex        =   6
      Top             =   5280
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
      TabIndex        =   11
      Top             =   2790
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   21430273
      CurrentDate     =   39889
   End
   Begin 合同管理.XPButton cmdCLSDate 
      Height          =   330
      Left            =   2835
      TabIndex        =   12
      Top             =   2760
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
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   1
      Left            =   1395
      TabIndex        =   0
      Top             =   2790
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
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   2
      Left            =   1395
      TabIndex        =   1
      Top             =   3360
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
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   3
      Left            =   1395
      TabIndex        =   2
      Top             =   3960
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
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   4
      Left            =   1395
      TabIndex        =   3
      Top             =   4560
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
      isNumber        =   -1  'True
      MaxLength       =   15
   End
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   5
      Left            =   5715
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2790
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
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   7
      Left            =   5715
      TabIndex        =   7
      Top             =   3960
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
   Begin 合同管理.FTextBox txtBorrow 
      Height          =   300
      Index           =   6
      Left            =   5715
      TabIndex        =   5
      Top             =   3360
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
      TabIndex        =   29
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "备    注"
      Height          =   160
      Left            =   4920
      TabIndex        =   28
      Top             =   3430
      Width           =   735
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
      TabIndex        =   27
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label29 
      Caption         =   "录入日期"
      Height          =   160
      Left            =   4875
      TabIndex        =   26
      Top             =   4020
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "借支人帐号"
      Height          =   160
      Left            =   4755
      TabIndex        =   25
      Top             =   2835
      Width           =   900
   End
   Begin VB.Label Label24 
      Caption         =   "借 支 人"
      Height          =   160
      Left            =   600
      TabIndex        =   24
      Top             =   3425
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "借支金额(元)"
      Height          =   165
      Left            =   240
      TabIndex        =   23
      Top             =   4020
      Width           =   1080
   End
   Begin VB.Label Label22 
      Caption         =   "借支余额(元)"
      Height          =   165
      Left            =   240
      TabIndex        =   22
      Top             =   4620
      Width           =   1080
   End
   Begin VB.Label Label21 
      Caption         =   "借支日期"
      Height          =   160
      Left            =   600
      TabIndex        =   21
      Top             =   2850
      Width           =   735
   End
   Begin VB.Label Label30 
      Caption         =   "承 揽 人"
      Height          =   160
      Left            =   600
      TabIndex        =   20
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblWTDW"
      Height          =   160
      Index           =   2
      Left            =   1560
      TabIndex        =   19
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label32 
      Caption         =   "合同名称"
      Height          =   160
      Left            =   5040
      TabIndex        =   18
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblHTMC"
      Height          =   160
      Index           =   3
      Left            =   6000
      TabIndex        =   17
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "合同编号"
      Height          =   160
      Left            =   600
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblHTBH"
      Height          =   160
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label35 
      Caption         =   "项目编号"
      Height          =   160
      Left            =   5040
      TabIndex        =   14
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblHTinfo 
      Caption         =   "lblXMBH"
      Height          =   160
      Index           =   1
      Left            =   6000
      TabIndex        =   13
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "frmInputBorrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dblBorrow As Double          ' 借支金额
Private Sub cmdCLS_Click(Index As Integer)
    txtBorrow(Index + 13).Text = ""
End Sub
Private Sub cmdCLSDate_Click()
    txtBorrow(1).Text = ""
End Sub
Private Sub cobCBFS_Expand()
    cmdSave.Enabled = True
End Sub

Private Sub DTPicker4_CloseUp()
    txtBorrow(1).Text = DTPicker4.value
End Sub

Private Sub Form_Activate()
    If UCase(DataOperateState) = "EDIT" Then
         txtBorrow(1).SetFocus
    End If

End Sub
Private Sub Form_Load()
'On Error GoTo aaaa
    'Me.WindowState = vbMaximized    '最大化窗口
    Width = 8700
    Height = 6480
    Me.Top = 500
    Me.Left = 1500
    imgIcon.Picture = frmMain.cmdLeft(2).Picture
    Label28.caption = ""
    txtBorrow(7).Text = Now()   '录入日期
    
    
    Select Case DataOperateState
        Case "EDIT"    '借支记录
            fillMain
            fillData
        
        Case "ADD"           '新增合同记录
            Me.caption = "数据录入"
            Label1.caption = Me.caption
            fillMain
            dblBalace = txtBorrow(4).Text
            
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
    strSQL = "select * from borrow where id=" & borrowID
    rs.Open strSQL, Conn, 1, 1
    
    dblBalace = dblBalace + rs("jzje")
    
    For i = 1 To rs.Fields.Count - 2
        If Not IsNull(rs.Fields(i).value) Then
            txtBorrow(i).Text = rs.Fields(i).value
        End If
    Next
    rs.Close
    Conn.Close
    
End Sub
Sub fillMain()
    Dim strSQL As String
    Dim i As Integer
    Dim balace As Double
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    DBConnect
    
    '获得总合同信息
    strSQL = "select htbh,xmbh,clr,htmc from main,sub where sub.id=" & subID & " and main.id=sub.zhtid"
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
    
    strSQL = "select ysjzje from sub where id=" & subID
    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Or IsNull(rs("ysjzje")) Then
        balace = 0
    Else
        balace = rs("ysjzje")
    End If
    
    rs.Close
    
    strSQL = "select sum(jzje) as jzzje from borrow where zhtid=" & subID
    rs.Open strSQL, Conn, 1, 1
    If Not IsNull(rs("jzzje")) Then
        balace = balace - rs("jzzje")
    End If
    
    txtBorrow(4).Text = balace
    
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
Private Sub txtBorrow_Change(Index As Integer)
    Label28.caption = ""
    If Index = 3 And Trim(txtBorrow(3).Text) <> "" Then
        If CDbl(Trim(txtBorrow(3).Text)) <> 0 Then
            txtBorrow(4).Text = dblBalace - CDbl(txtBorrow(3).Text)
        End If
    End If
    cmdSave.Enabled = True
    
End Sub
Private Sub txtBorrow_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Index < 5 Then
        txtBorrow(Index + 1).SetFocus
    End If
End Sub
Private Sub cmdSave_Click()
   ' On Error GoTo errmsg
    
    If Trim(txtBorrow(1).Text) = "" Then
        MsgBox "未填写[借支日期]", vbInformation, Me.caption
        txtBorrow(1).SetFocus
        Exit Sub
    End If
    
    Select Case DataOperateState
        Case "EDIT"
            UpdateData "borrow", 7, borrowID
            
        Case "ADD"
            AddRecord "borrow", 7
            
            dblBalace = txtBorrow(4).Text
            For i = 1 To 6
                txtBorrow(i).Text = ""
            Next
            
            txtBorrow(4).Text = dblBalace
            
            txtBorrow(7).Text = Now()
            cmdSave.Enabled = False
            Label28.caption = "保存成功!"
            txtBorrow(1).SetFocus
            
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
                    
        strText = Trim(txtBorrow(i).Text)
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
        strText = Trim(txtBorrow(i).Text)
                                    
        If strText <> "" Then
            strField = strField & rs.Fields(i).Name & ","
            If FieldTypeIsChar(rs.Fields(i).Type) Then   '字符或日期字段,加单引号
                strText = "'" & strText & "'"
            End If
            strValue = strValue & strText & ","
         End If
    
    Next
    strField = strField & "zhtid)"
    strValue = strValue & subID & ")"
    
    strSQL = "insert into " & strTable & " " & strField & "  values" & strValue
    'MsgBox strSQL
    Conn.Execute strSQL

End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

