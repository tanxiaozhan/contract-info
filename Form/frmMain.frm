VERSION 5.00
Object = "{A4BF9E9F-333F-4D07-A80E-DA359D576BFF}#3.0#0"; "xpmenu.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "��ͬ����"
   ClientHeight    =   9210
   ClientLeft      =   285
   ClientTop       =   705
   ClientWidth     =   11400
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":1C7A
   StartUpPosition =   2  '��Ļ����
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
         Caption         =   "��ӭʹ�ñ�ϵͳ"
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
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "�����ѯ"
         ToolTip         =   "���б�ʽ�������ѯ��ͬ��Ϣ"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2970D
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdClose 
         Height          =   195
         Left            =   1245
         TabIndex        =   3
         Top             =   60
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   344
         Caption         =   "��"
         ToolTip         =   "�ر�"
         BackColor       =   6956042
         ForeColor       =   16777215
         MouseDownColor  =   6956042
         MouseOnColor    =   6956042
         StyleColor      =   0
         Style3dColor1   =   16577259
         Style3dColor2   =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "���㵥"
         ToolTip         =   "���ɽ��㵥"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2A3E7
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   3540
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "ȷ�ϵ�"
         ToolTip         =   "����ȷ�ϵ�"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2B0C1
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1500
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "����¼��"
         ToolTip         =   "¼���ܺ�ͬ��Ϣ"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2BD9B
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   5
         Left            =   225
         TabIndex        =   9
         Top             =   4560
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "��֧��"
         ToolTip         =   "���ɽ�֧��"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2CA75
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   6
         Left            =   225
         TabIndex        =   10
         Top             =   5580
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "�û�����"
         ToolTip         =   "��¼�û���Ϣ����"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2D74F
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdLeft 
         Height          =   885
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   6600
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1561
         Caption         =   "ѡ������"
         ToolTip         =   "���ù�����С�����"
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2E429
         style           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "������"
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
      Begin ��ͬ����.XButton tbLogin 
         Height          =   330
         Left            =   210
         TabIndex        =   13
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Caption         =   ""
         ToolTip         =   "���ص�½����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2F46D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   2
         Left            =   1530
         TabIndex        =   14
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "¼��"
         ToolTip         =   "¼���ܺ�ͬ��Ϣ"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2FA07
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   3
         Left            =   2460
         TabIndex        =   15
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "���㵥"
         ToolTip         =   "���ɽ��㵥"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":2FFA1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   16
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "���"
         ToolTip         =   "���б�ʽ�������ѯ��ͬ��Ϣ"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3053B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   4
         Left            =   3390
         TabIndex        =   17
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "ȷ�ϵ�"
         ToolTip         =   "����ȷ�ϵ�"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":30AD5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   5
         Left            =   4320
         TabIndex        =   18
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "��֧��"
         ToolTip         =   "���ɽ�֧��"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3106F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   6
         Left            =   5250
         TabIndex        =   19
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "�û�"
         ToolTip         =   "��¼�û���Ϣ����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":31609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton cmdAbout 
         Height          =   330
         Left            =   7500
         TabIndex        =   20
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Caption         =   ""
         ToolTip         =   "���ڱ����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":31BA3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbLeft 
         Height          =   330
         Index           =   7
         Left            =   6120
         TabIndex        =   22
         Top             =   30
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   582
         Caption         =   "ѡ��"
         ToolTip         =   "��¼�û���Ϣ����"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":3213D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ��ͬ����.XButton tbExit 
         Height          =   330
         Left            =   8010
         TabIndex        =   23
         Top             =   30
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         Caption         =   ""
         ToolTip         =   "�˳�������"
         BackColor       =   13752539
         MouseDownColor  =   12363422
         MouseOnColor    =   14204854
         StyleColor      =   6956042
         Style3dColor1   =   6956042
         Style3dColor2   =   6956042
         Picture         =   "frmMain.frx":326D7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuExItem 
         Caption         =   "������Ŀ���ϣ��ɣ�(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuExItemNew 
         Caption         =   "������Ŀ���ϣ��£�"
      End
      Begin VB.Menu mnuExIncome 
         Caption         =   "�����տ�һ����(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuContractList 
         Caption         =   "������̨ͬ��"
      End
      Begin VB.Menu mnuFileSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDBBackUp 
         Caption         =   "�������ݿ�"
      End
      Begin VB.Menu mnuDBResume 
         Caption         =   "�ָ����ݿ�"
      End
      Begin VB.Menu mnuFileSP2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogin 
         Caption         =   "���ص�½����(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFileSP3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "��ͼ(&V)"
      Begin VB.Menu mnuLeft 
         Caption         =   "�����ѯ(&B)"
         Index           =   1
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "����¼��(&D)"
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "���㵥(&J)"
         Index           =   3
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "ȷ�ϵ�(&C)"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "��֧��(&W)"
         Index           =   5
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "�û�����(&U)"
         Index           =   6
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "ѡ������(&O)"
         Index           =   7
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuViewSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "������(&G)"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuTB 
         Caption         =   "������(&T)"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSB 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuContent 
         Caption         =   "����(&C)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSupply 
         Caption         =   "����֧��(&S)"
      End
      Begin VB.Menu mnuHelpSP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "���ڱ����(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�϶������API
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
    SetSB 1, "����λ�ã�" & cmdLeft(Index).caption
    
    curDOCType = Index - 2  '�ĵ����ͣ�1-���㵥��2-��Ŀȷ�ϵ���3-��Ŀ��֧��
    
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
    
    
    '��ȡ����λ��,��ͼ��Ϣ
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
    '�ж��û�����,1-����Ա����2-��ͨ�û���ֻ�ܲ鿴��
    cmdLeft(6).Enabled = (curUserLevel = 1)
    tbLeft(6).Enabled = (curUserLevel = 1)
    mnuLeft(6).Enabled = (curUserLevel = 1)
    LastFrm = 0
    
    '�����б����
    
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
        strSQL = "alter table main add fzr CHAR(30) WITH COMP"   'unicodeѹ�����ı����ֶ�
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
    MsgBox "��ͬ������� V1.0" & Chr(13) & Chr(13) & "    2009.03", vbInformation
End Sub

Private Sub mnuContent_Click()
    MsgBox "���ް���������£�", vbInformation
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
    
    Dlg.Filter = "��ͬ���������ļ�(*.htb)|*.htb"
    Dlg.FileName = "DATA" & Format(Now(), "yyyy-mm-dd hh.mm.ss") & ".htb"
    Dlg.DialogTitle = "���ݱ���"
    Dlg.InitDir = GetApp & "bak"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    FileCopy GetApp & "data.htb", Dlg.FileName
    MsgBox "���ݱ��ݳɹ���", vbInformation, "���ݱ���"
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "���ݱ���"
End Sub

Private Sub mnuDBResume_Click()
    On Error GoTo errmsg
    
    If Conn.State <> 0 Then
        Conn.Close
    End If
    If DirExists(GetApp & "bak") <> 0 Then
        Dlg.InitDir = GetApp & "bak"
    End If
    
    Dlg.Filter = "��ͬ���������ļ�(*.htb)|*.htb"
    Dlg.DialogTitle = "���ݻָ�"
    Dlg.CancelError = True
    Dlg.ShowOpen
    
    If MsgBox("���棺���ݻָ�����" & Dlg.FileName & "�����ݸ������������ݡ�", vbExclamation + vbYesNo, "���ݻָ�") = vbNo Then Exit Sub
    If MsgBox("ȷ�Ͻ������ݻָ���?", vbExclamation + vbYesNo, "���ݻָ�") = vbNo Then Exit Sub
    FileCopy Dlg.FileName, GetApp & "data.htb"
    MsgBox "���ݻָ��ɹ���", vbInformation, "���ݻָ�"
    
    cmdLeft_Click 1              '�����б���
    frmList.loadList
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "���ݻָ�"

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
    Dim strHTBH, strXMBH As String '��ͬ���,��Ŀ���
    Dim dblBalace As Double    '��֧���
    
    startRow = 3  '�ӵ�3�п�ʼ���
    
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
        MsgBox "δ�ҵ���ؼ�¼��������ֹ��", vbExclamation, "������Ŀ����"
        rs.Close
        Conn.Close
        Exit Sub
    End If
    
    Dlg.Filter = "MS Excel�ļ�(*.xls)|*.xls"
    Dlg.FileName = "��Ŀ����(" & Format(Now(), "yyyy-mm-dd") & ")"
    Dlg.DialogTitle = "������Ŀ����"
    Dlg.InitDir = GetApp & "Doc"
    Dlg.CancelError = True
    Dlg.ShowSave
    
    strFormat = ";;;;;;;yyyy��mm��dd��;yyyy��mm��dd��;##,##0.00;yyyy��mm��dd��;##,##0.00;##,##0.00;##,##0.00;yyyy��mm��dd��"
    arrayFormat = Split(strFormat, ";")
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(GetApp & "templets\��Ŀ����.xls")
    xlApp.Visible = False
    Set xlSheet = xlBook.Worksheets("Sheet1")
    
    strXMBH = ""    '��Ŀ���
    strHTBH = ""   '��ͬ���
    n = 0
    row = 1
    
    
    
    Do While Not rs.EOF
        n = n + 1
        
        xlSheet.Cells(startRow + row, 1) = Trim(CStr(n))  '��4�У�1��
        xlSheet.Cells(startRow + row, 2) = IIf(rs("yjs"), "��", "��") '��4�У�2��
        If rs("yjs") Then xlSheet.Cells(startRow + row, 2).Font.ColorIndex = 3
        
        If IsNull(rs("ysjzje")) Then      'Ԥ���֧���
            dblBalace = 0
        Else
            dblBalace = CDbl(rs("ysjzje"))
        End If
            
        For i = 1 To 9 '1-��Ŀ���,....9-Ԥ���֧���
            If Not IsNull(rs.Fields(i).value) Then
                xlSheet.Cells(startRow + row, 2 + i) = IIf(arrayFormat(i) <> "", Format(CStr(rs.Fields(i).value), arrayFormat(i)), rs.Fields(i).value)
                    
            End If
        Next
        
        For i = 10 To 11   '10-�����,11-��������
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
            
                For i = 0 To 1    '��֧���
                    If Not IsNull(rsBorrow.Fields(i).value) Then
                        xlSheet.Cells(startRow + row, 12 + i) = IIf(arrayFormat(10 + i) <> "", Format(CStr(rsBorrow.Fields(i).value), arrayFormat(10 + i)), rsBorrow.Fields(i).value)
                    End If
                
                Next
            
                If Not IsNull(rsBorrow("jzje")) Then    '�����֧���
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
    
    MsgBox "��Ŀ���ϵ�����ɣ�" & Chr(13) & "���浽" & Dlg.FileName, vbInformation, "������Ŀ����"
    
    Exit Sub

errmsg:
    If Err.Number = 32755 Then Exit Sub     '32755���û����ȡ����ť
    MsgBox Err.Description, vbInformation, "������Ŀ����"

End Sub
Private Sub mnuExItemNew_Click()    '������Ŀ���ϣ��£�
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
    MsgBox "���µ磺31304837", vbInformation
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


