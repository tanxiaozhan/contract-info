VERSION 5.00
Object = "{A4BF9E9F-333F-4D07-A80E-DA359D576BFF}#3.0#0"; "xpmenu.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "合同管理"
   ClientHeight    =   9210
   ClientLeft      =   285
   ClientTop       =   705
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":1C7A
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   2760
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin HookMenu.XPMenu XPMenu1 
      Left            =   4200
      Top             =   3120
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UserSelectedMenuBackColour=   13040639
      UserSelectedMenuBorderColour=   16711680
      UserTopMenuBackColour=   16761765
      UserTopMenuSelectedColour=   16769990
      UserTopMenuHotColour=   33023
      UserTopMenuHotBorderColour=   16711680
      UserMenuBorderColour=   8388608
      UserCheckBackColour=   8108783
      UserCheckBorderColour=   16711680
      UserGradientOne =   16777215
      UserGradientTwo =   16761765
      UserUseGradient =   -1  'True
      UserSelectedItemForeColour=   16384
      UserSideBarColour=   16711680
      CopyRight       =   "lcXPMenu"
   End
   Begin VB.PictureBox picSB 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   760
      TabIndex        =   4
      Top             =   8910
      Width           =   11400
      Begin VB.Image Image2 
         Height          =   240
         Left            =   75
         Picture         =   "frmMain.frx":288EF
         Top             =   45
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3150
         Picture         =   "frmMain.frx":28E79
         Top             =   45
         Width           =   240
      End
      Begin VB.Label LBSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   180
         Index           =   2
         Left            =   3465
         TabIndex        =   12
         Top             =   75
         Width           =   90
      End
      Begin VB.Label LBSB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "欢迎使用本系统"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   11
         Top             =   75
         Width           =   1260
      End
      Begin VB.Shape Shb2 
         BorderColor     =   &H00A6A6A6&
         Height          =   270
         Left            =   3090
         Top             =   30
         Width           =   6885
      End
      Begin VB.Image imgLB 
         Height          =   180
         Left            =   10080
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmMain.frx":29203
         Top             =   120
         Width           =   180
      End
      Begin VB.Shape Shb1 
         BorderColor     =   &H00A6A6A6&
         Height          =   270
         Left            =   30
         Top             =   30
         Width           =   3015
      End
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   0
      ScaleHeight     =   568
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   1
      Top             =   390
      Width           =   1560
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "浏览查询"
         ToolTip         =   "按列表方式浏览、查询合同信息"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2970D
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdClose 
         Height          =   195
         Left            =   1245
         TabIndex        =   3
         Top             =   60
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   344
         Caption         =   "×"
         ToolTip         =   "关闭"
         BackColor       =   6956042
         ForeColor       =   16777215
         MouseDownColor  =   6956042
         MouseOnColor    =   6956042
         StyleColor      =   0
         Style3dColor1   =   16577259
         Style3dColor2   =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "结算单"
         ToolTip         =   "生成结算单"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2A3E7
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   3540
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "确认单"
         ToolTip         =   "生成确认单"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2B0C1
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1500
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "数据录入"
         ToolTip         =   "录入总合同信息"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2BD9B
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   5
         Left            =   225
         TabIndex        =   9
         Top             =   4560
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "借支单"
         ToolTip         =   "生成借支单"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2CA75
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   6
         Left            =   225
         TabIndex        =   10
         Top             =   5580
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "用户管理"
         ToolTip         =   "登录用户信息管理"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2D74F
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdLeft 
         Height          =   885
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   6600
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "选项设置"
         ToolTip         =   "设置工作量小数宽度"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2E429
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape ShLeft 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00A6A6A6&
         Height          =   7575
         Left            =   30
         Top             =   330
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "导航栏"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   90
         TabIndex        =   2
         Top             =   75
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H006A240A&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   270
         Left            =   30
         Top             =   30
         Width           =   1485
      End
   End
   Begin VB.PictureBox picTB 
      Align           =   1  'Align Top
      BackColor       =   &H00D1D8DB&
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   0
      Picture         =   "frmMain.frx":2ED03
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   760
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      Begin 合同管理.XButton tbLogin 
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Caption         =   ""
         ToolTip         =   "返回登陆窗口"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2F46D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   2
         Left            =   1530
         TabIndex        =   14
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "录入"
         ToolTip         =   "录入总合同信息"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2FA07
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   3
         Left            =   2460
         TabIndex        =   15
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "结算单"
         ToolTip         =   "生成结算单"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2FFA1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "浏览"
         ToolTip         =   "按列表方式浏览、查询合同信息"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3053B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   4
         Left            =   3390
         TabIndex        =   17
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "确认单"
         ToolTip         =   "生成确认单"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":30AD5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   5
         Left            =   4320
         TabIndex        =   18
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "借支单"
         ToolTip         =   "生成借支单"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3106F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   6
         Left            =   5250
         TabIndex        =   19
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "用户"
         ToolTip         =   "登录用户信息管理"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":31609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton cmdAbout 
         Height          =   330
         Left            =   7500
         TabIndex        =   20
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Caption         =   ""
         ToolTip         =   "关于本软件"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":31BA3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbLeft 
         Height          =   330
         Index           =   7
         Left            =   6120
         TabIndex        =   22
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "选项"
         ToolTip         =   "登录用户信息管理"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3213D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin 合同管理.XButton tbExit 
         Height          =   330
         Left            =   8010
         TabIndex        =   23
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Caption         =   ""
         ToolTip         =   "退出本程序"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":326D7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00A6A6A6&
         X1              =   494
         X2              =   494
         Y1              =   3
         Y2              =   23
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00A6A6A6&
         X1              =   42
         X2              =   42
         Y1              =   3
         Y2              =   23
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuExItem 
         Caption         =   "导出项目资料（旧）(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuExItemNew 
         Caption         =   "导出项目资料（新）"
      End
      Begin VB.Menu mnuExIncome 
         Caption         =   "导出收款一览表(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuContractList 
         Caption         =   "导出合同台帐"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBBackUp 
         Caption         =   "备份数据库"
      End
      Begin VB.Menu mnuDBResume 
         Caption         =   "恢复数据库"
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "返回登陆界面(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "视图(&V)"
      Begin VB.Menu mnuLeft 
         Caption         =   "浏览查询(&B)"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "数据录入(&D)"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "结算单(&J)"
         Index           =   3
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "确认单(&C)"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "借支单(&W)"
         Index           =   5
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "用户管理(&U)"
         Index           =   6
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "选项设置(&O)"
         Index           =   7
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuViewSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "导航栏(&G)"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuTB 
         Caption         =   "工具条(&T)"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSB 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuContent 
         Caption         =   "内容(&C)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSupply 
         Caption         =   "技术支持(&S)"
      End
      Begin VB.Menu mnuHelpSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "关于本软件(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'拖动窗体的API
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim CanResize As Boolean
Public LastFrm As Long

Private Sub cmdAbout_Click()
    mnuAbout_Click
End Sub

Private Sub cmdClose_Click()
    picLeft.Visible = False
    mnuGuide.Checked = False
    SaveINI "Main", "Guide", "n"
End Sub

Public Sub cmdLeft_Click(Index As Integer)
    If LastFrm = Index Then Exit Sub
    If LastFrm > 0 Then
        cmdLeft(LastFrm).IfDraw = False
        tbLeft(LastFrm).IfDraw = False
        mnuLeft(LastFrm).Checked = False
        cmdLeft(LastFrm).BackColor = picLeft.BackColor
        tbLeft(LastFrm).BackColor = picTB.BackColor
    Else
        'Unload frmList
    End If
    Select Case LastFrm
        Case 1: Unload frmList
        Case 2: Unload frmInputMain
        Case 3, 4, 5: Unload frmDoc
        Case 6: Unload frmUser
        Case 7: Unload frmOption
    End Select

    LastFrm = Index
    cmdLeft(Index).IfDraw = True
    tbLeft(Index).IfDraw = True
    mnuLeft(Index).Checked = True
    cmdLeft(Index).BackColor = 14210516
    tbLeft(Index).BackColor = 14210516
    SetSB 1, "现在位置：" & cmdLeft(Index).caption
    
    curDOCType = Index - 2  '文档类型：1-结算单，2-项目确认单，3-项目借支单
    
    Select Case Index
        Case 1: frmList.Show
        Case 2: frmInputMain.Show
        Case 3, 4, 5: frmDoc.Show
        Case 6: frmUser.Show
        Case 7: frmOption.Show
    End Select
End Sub

Private Sub imgLB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(hwnd, &HA1, 17, 0)
    End If
End Sub

Private Sub imgLogin_Click()

End Sub

Private Sub MDIForm_Load()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    
    '读取窗体位置,视图信息
    If GetINI("Main", "Left") = "" Then
        Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    Else
        Move GetLongINI("Main", "Left"), GetLongINI("Main", "Top"), GetLongINI("Main", "Width"), GetLongINI("Main", "Height")
        Dim j As Long
        j = GetLongINI("Main", "WindowState")
        If j = 2 Then Me.WindowState = 2
    End If
    CanResize = True
    If GetINI("Main", "Guide") = "n" Then
        picLeft.Visible = False
        mnuGuide.Checked = False
    End If
    If GetINI("Main", "ToolBar") = "n" Then
        picTB.Visible = False
        mnuTB.Checked = False
    End If
    If GetINI("Main", "StateBar") = "n" Then
        picSB.Visible = False
        mnuSB.Checked = False
    End If
    '判断用户类型,1-管理员级别，2-普通用户（只能查看）
    cmdLeft(6).Enabled = (curUserLevel = 1)
    tbLeft(6).Enabled = (curUserLevel = 1)
    mnuLeft(6).Enabled = (curUserLevel = 1)
    LastFrm = 0
    
    '加载列表参数
    
    DBConnect
    strSQL = "select * from options"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        curList1Index = 0
        curList2Index = 0
        curList3Index = 0
        curList4Index = 0
        curList5Index = 0
    Else
        curList1Index = rs("List1Index")
        curList2Index = rs("List2Index")
        curList3Index = rs("List3Index")
        curList4Index = rs("List4Index")
        curList5Index = rs("List5Index")
    End If
    rs.Close
    
    strSQL = "select top 1 * from main"
    rs.Open strSQL, Conn, 1, 1
    
    fieldCount = rs.Fields.Count
    For i = 0 To fieldCount - 1
        If Trim(rs.Fields(i).Name) = "fzr" Then Exit For
    Next
    
    rs.Close
    Set rs = Nothing
    
    If i >= fieldCount Then
        strSQL = "alter table main add fzr CHAR(30) WITH COMP"   'unicode压缩的文本型字段
        Conn.Execute strSQL
    End If
    
    Conn.Close
    
    Me.cmdLeft_Click 1
    
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    If CanResize = False Then Exit Sub
    If Me.Width < 9900 Then Me.Width = 9900
    If Me.Height < 8370 Then Me.Height = 8370
    SaveINI "Main", "WindowState", CStr(WindowState)
    If Me.WindowState = 0 Then
        SaveINI "Main", "Width", CStr(Width)
        SaveINI "Main", "Height", CStr(Height)
    End If
    picSB_Resize
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
    DBConnect
    Conn.Execute "update options set list1Index=" & curList1Index & ",list2index=" & curList2Index & ",list3index=" & curList3Index & ",List5index=" & curList5Index
    Conn.Close
    Set frmMain = Nothing
End Sub

Private Sub mnuAbout_Click()
    MsgBox "合同管理程序 V1.0" & Chr(13) & Chr(13) & "    2009.03", vbInformation
End Sub

Private Sub mnuContent_Click()
    MsgBox "暂无帮助，请见谅！", vbInformation
End Sub

Private Sub mnuContractList_Click()
    frmInputContractNo.Show vbModal
End Sub

Private Sub mnuDBBackUp_Click()
    On Error GoTo errmsg
    
    If Conn.State <> 0 Then
        Conn.Close
    End If
    
    If DirExists(GetApp & "bak") = 0 Then
        MkDir GetApp & "bak"
    End If
    
    Dlg.Filter = "合同管理数据文件(*.htb)|*.htb"
    Dlg.FileName = "DATA" & Format(Now(), "yyyy-mm-dd hh.mm.ss") & ".htb"
    Dlg.DialogTitle = "数据备份"
    Dlg.InitDir = GetApp & "bak"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    FileCopy GetApp & "data.htb", Dlg.FileName
    MsgBox "数据备份成功！", vbInformation, "数据备份"
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755，用户点击取消按钮
    MsgBox Err.Description, vbInformation, "数据备份"
End Sub

Private Sub mnuDBResume_Click()
    On Error GoTo errmsg
    
    If Conn.State <> 0 Then
        Conn.Close
    End If
    If DirExists(GetApp & "bak") <> 0 Then
        Dlg.InitDir = GetApp & "bak"
    End If
    
    Dlg.Filter = "合同管理数据文件(*.htb)|*.htb"
    Dlg.DialogTitle = "数据恢复"
    Dlg.CancelError = True
    Dlg.ShowOpen
    
    If MsgBox("警告：数据恢复将用" & Dlg.FileName & "的数据覆盖现在有数据。", vbExclamation + vbYesNo, "数据恢复") = vbNo Then Exit Sub
    If MsgBox("确认进行数据恢复吗?", vbExclamation + vbYesNo, "数据恢复") = vbNo Then Exit Sub
    FileCopy Dlg.FileName, GetApp & "data.htb"
    MsgBox "数据恢复成功！", vbInformation, "数据恢复"
    
    cmdLeft_Click 1              '加载列表窗口
    frmList.loadList
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755，用户点击取消按钮
    MsgBox Err.Description, vbInformation, "数据恢复"

End Sub

Private Sub mnuExIncome_Click()
    frmExportIncomeYear.Show vbModal
End Sub

Private Sub mnuExit_Click()
    frmList.SaveListColWidth
    Unload Me
End Sub

Private Sub mnuExItem_Click()
    On Error GoTo errmsg
    Dim xlApp As excel.Application
    Dim xlBook As excel.Workbook
    Dim xlSheet As excel.Worksheet
    Dim xlRange As excel.Range
    Dim rs, rsBorrow As ADODB.Recordset
    Dim strSQL As String
    Dim i, row, startRow, n As Integer
    Dim strFormat As String
    Dim strHTBH, strXMBH As String '合同编号,项目编号
    Dim dblBalace As Double    '借支余额
    
    startRow = 3  '从第3行开始填充
    
    Set rs = New ADODB.Recordset
    Set rsBorrow = New ADODB.Recordset
    DBConnect
    
    If DirExists(GetApp & "Doc") = 0 Then
        MkDir GetApp & "Doc"
    End If
    
    strSQL = "select  sub.yjs,sub.xmbh,main.wtdw,main.wtdwlxr,main.wtdwlxdh,sub.xmmc,sub.clr," & _
                  "sub.jcrq,sub.tcrq,sub.ysjzje,sub.jsj,sub.jsrq,sub.id" & " " & _
             "from main,sub" & " " & _
             "where main.id=sub.zhtid" & " " & _
             "order by main.lrrq desc,sub.xmbh"
    

    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        MsgBox "未找到相关记录，导出中止！", vbExclamation, "导出项目资料"
        rs.Close
        Conn.Close
        Exit Sub
    End If
    
    Dlg.Filter = "MS Excel文件(*.xls)|*.xls"
    Dlg.FileName = "项目资料(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "导出项目资料"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    strFormat = ";;;;;;;yyyy年mm月dd日;yyyy年mm月dd日;##,##0.00;yyyy年mm月dd日;##,##0.00;##,##0.00;##,##0.00;yyyy年mm月dd日"
    arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\项目资料.xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    strXMBH = ""    '项目编号
    strHTBH = ""   '合同编号
    n = 0
    row = 1
    
    
    
    Do While Not rs.EOF
        n = n + 1
        
        xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '第4行，1列
        xlSheet.Cells(startRow + row, 2) = IIf(rs("yjs"), "是", "否") '第4行，2列
        If rs("yjs") Then xlSheet.Cells(startRow + row, 2).Font.ColorIndex = 3
        
        If IsNull(rs("ysjzje")) Then      '预算借支金额
            dblBalace = 0
        Else
            dblBalace = CDbl(rs("ysjzje"))
        End If
            
        For i = 1 To 9 '1-项目编号,....9-预算借支金额
            If Not IsNull(rs.Fields(i).value) Then
                xlSheet.Cells(startRow + row, 2 + i) = IIf(arrayFormat(i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(i)), rs.Fields(i).value)
                    
            End If
        Next
        
        For i = 10 To 11   '10-结算价,11-结算日期
            If Not IsNull(rs.Fields(i).value) Then
                xlSheet.Cells(startRow + row, 5 + i) = IIf(arrayFormat(3 + i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(3 + i)), rs.Fields(i).value)
            End If
        Next
        
            
        strSQL = "select jzrq,jzje from borrow where zhtid=" & rs("id") & " order by jzrq"
        rsBorrow.Open strSQL, Conn, 1, 1
            
    
        If rsBorrow.RecordCount < 1 Then
            row = row + 1
        Else
        
            
            Do While Not rsBorrow.EOF
            
                For i = 0 To 1    '借支情况
                    If Not IsNull(rsBorrow.Fields(i).value) Then
                        xlSheet.Cells(startRow + row, 12 + i) = IIf(arrayFormat(10 + i) <> "", Format(CStr(rsBorrow.Fields(i).value), arrayFormat(10 + i)), rsBorrow.Fields(i).value)
                    End If
                
                Next
            
                If Not IsNull(rsBorrow("jzje")) Then    '计算借支余额
                    dblBalace = dblBalace - CDbl(rsBorrow("jzje"))
                End If
                xlSheet.Cells(startRow + row, 14) = IIf(arrayFormat(12) <> "", Format(CStr(dblBalace), arrayFormat(12)), CStr(dblBalace))
                
                rsBorrow.MoveNext
                row = row + 1
            Loop
            
            If rsBorrow.RecordCount > 1 Then
                For i = 1 To 11
                    xlSheet.Range(xlSheet.Cells(startRow + row - 1, i), xlSheet.Cells(startRow + row - rsBorrow.RecordCount, i)).Merge
                Next
                For i = 15 To 16
                    xlSheet.Range(xlSheet.Cells(startRow + row - 1, i), xlSheet.Cells(startRow + row - rsBorrow.RecordCount, i)).Merge
                Next
            
            End If
        
        End If
        
        rsBorrow.Close
        
        rs.MoveNext
    Loop
    
    Set xlRange = xlSheet.Range(xlSheet.Cells(startRow, 1), xlSheet.Cells(startRow + row - 1, 16))
    
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

    
    
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing
    xlBook.SaveAs Dlg.FileName
    xlBook.Close (True)
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "项目资料导出完成！" & Chr(13) & "保存到" & Dlg.FileName, vbInformation, "导出项目资料"
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755，用户点击取消按钮
    MsgBox Err.Description, vbInformation, "导出项目资料"

End Sub
Private Sub mnuExItemNew_Click()    '导出项目资料（新）
    frmExportItem.Show vbModal
End Sub

Private Sub mnuGuide_Click()
    mnuGuide.Checked = Not mnuGuide.Checked
    picLeft.Visible = mnuGuide.Checked
    SaveINI "Main", "Guide", IIf(mnuGuide.Checked = True, "", "n")
End Sub

Private Sub mnuLeft_Click(Index As Integer)
    cmdLeft_Click Index
End Sub

Private Sub mnuLogin_Click()
On Error Resume Next
    Unload Me
    frmLogin.Show
End Sub

Private Sub mnuSupply_Click()
    MsgBox "请致电：31304837", vbInformation
End Sub

Private Sub picSB_Resize()
On Error Resume Next
    Shb2.Width = Me.Width / 15 - IIf(Me.WindowState = 2, 210, 230)
    imgLB.Visible = (Me.WindowState <> 2)
    imgLB.Left = Me.Width / 15 - 20
End Sub

Private Sub mnuSB_Click()
    mnuSB.Checked = Not mnuSB.Checked
    picSB.Visible = mnuSB.Checked
    SaveINI "Main", "StateBar", IIf(mnuSB.Checked = True, "", "n")
End Sub

Private Sub mnuTB_Click()
    mnuTB.Checked = Not mnuTB.Checked
    picTB.Visible = mnuTB.Checked
    SaveINI "Main", "ToolBar", IIf(mnuTB.Checked = True, "", "n")
End Sub

Private Sub picLeft_Resize()
On Error Resume Next
    ShLeft.Height = picLeft.Height / 15 - 23
End Sub

Private Sub tbExit_Click()
    mnuExit_Click
End Sub

Private Sub tbLeft_Click(Index As Integer)
    cmdLeft_Click Index
End Sub

Private Sub tbLogin_Click()
    mnuLogin_Click
End Sub


