VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInputContractNo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������̨ͬ��"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4500
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   420
      Left            =   315
      TabIndex        =   5
      Top             =   780
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
   Begin ��ͬ����.FTextBox txtyear 
      Height          =   300
      Left            =   2340
      TabIndex        =   2
      Top             =   825
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      BackColor       =   12636398
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "����"
      FontSize        =   9
      isNumber        =   -1  'True
      MaxLength       =   5
      afterdecimal    =   0
   End
   Begin ��ͬ����.XPButton cmdExport 
      Height          =   450
      Left            =   450
      TabIndex        =   1
      Top             =   1890
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   794
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
   Begin ��ͬ����.XPButton XPButton1 
      Height          =   450
      Left            =   2400
      TabIndex        =   0
      Top             =   1890
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   794
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
   Begin VB.Label Label3 
      Caption         =   "��"
      Height          =   240
      Left            =   3150
      TabIndex        =   7
      Top             =   885
      Width           =   420
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "���ڵ�����̨ͬ�ʣ����Ժ�..."
      BeginProperty Font 
         Name            =   "����"
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
   Begin VB.Label Label2 
      Caption         =   "��ʾ����������ݿɵ��������������"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   705
      TabIndex        =   4
      Top             =   1455
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "���������(4λ)"
      Height          =   270
      Left            =   780
      TabIndex        =   3
      Top             =   900
      Width           =   1425
   End
End
Attribute VB_Name = "frmInputContractNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    txtyear.SetFocus
    lblInfo.Visible = False
    pbar.Visible = False

End Sub

Private Sub XPButton1_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
On Error GoTo errmsg
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim strYear1, strYear2 As String '��������ֹ���
    Dim strTitle As String
    
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range
        
    Label1.Visible = False
    txtyear.Visible = False
    Label2.Visible = False
    
    startRow = 3  '�ӵ�3�п�ʼ���
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    Dlg.Filter = "MS Excel�ļ�(*.xls)|*.xls"
    Dlg.FileName = "��̨ͬ��(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "������̨ͬ��"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    pbar.Visible = True
    lblInfo.Visible = True
    
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\��̨ͬ��.xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    n = 0
    row = 0
    
    sql = "select main.id as mainid,main.htbh,main.wtdw,main.wtdwlxr,main.wtdwlxdh,main.htmc,main.htzj,main.gzny " & _
          "from main "
        
    If Trim(txtyear.Text) <> "" Then
        sql = sql & " " & "where htbh like '" & Trim(txtyear.Text) & "%'"
    End If
    
    sql = sql & " " & "order by main.htbh"
    
    DBConnect
    Set rs = New ADODB.Recordset
    
    rs.Open sql, Conn, 1, 1
    
    pbar.Max = rs.RecordCount
    strYear1 = ""   '��ʼ���
    strYear2 = ""   '�������
    If rs.RecordCount > 0 Then strYear1 = Left(Trim(rs("htbh")), 4)
    Do While Not rs.EOF
        row = row + 1
        n = n + 1
        xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))
        xlSheet.Cells(startRow + row, 3) = Trim(rs("htbh"))
        xlSheet.Cells(startRow + row, 4) = rs("wtdw")
        xlSheet.Cells(startRow + row, 5) = rs("wtdwlxr")
        xlSheet.Cells(startRow + row, 6) = rs("wtdwlxdh")
        xlSheet.Cells(startRow + row, 7) = rs("htmc")
        xlSheet.Cells(startRow + row, 8) = rs("gzny")
        xlSheet.Cells(startRow + row, 9) = Format(rs("htzj"), "##,##0.00")
        
        strYear2 = Left(Trim(rs("htbh")), 4)
        
        rs.MoveNext
        pbar.value = pbar.value + 1
    Loop
    rs.Close
    
    'Excel�ļ����⴦��
    strTitle = ""
    If strYear1 <> "" Then strTitle = "(" & strYear1 & "��)"
    If strYear1 <> strYear2 And strYear1 <> "" And strYear2 <> "" Then strTitle = "(" & strYear1 & "--" & strYear2 & "��)"
    strTitle = "��̨ͬ��" & strTitle
    xlSheet.Cells(1, 1) = strTitle
    
    Set xlRange = xlSheet.Range(xlSheet.Cells(startRow, 1), xlSheet.Cells(startRow + row, 11))
    
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
    
    
    xlSheet.Cells(2, 1).Select
    
    
    xlApp.DisplayAlerts = False   '���治��ʾ������ʾ
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    pbar.value = pbar.Max
    
    MsgBox "��̨ͬ�ʵ�����ɣ�" & Chr(13) & "���浽" & Dlg.FileName, vbInformation, "������̨ͬ��"
    
    pbar.Visible = False
    lblInfo.Visible = False
    Label2.Visible = True
    Label1.Visible = True
    txtyear.Visible = True
        
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
