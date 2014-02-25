VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmList 
   AutoRedraw      =   -1  'True
   Caption         =   "合同列表"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   ControlBox      =   0   'False
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   518
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   623
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MG2 
      Height          =   195
      Left            =   1365
      TabIndex        =   2
      Top             =   555
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   344
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSComctlLib.ListView List6 
      Height          =   495
      Left            =   2370
      TabIndex        =   30
      Top             =   510
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "分项内容"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "工作量(KM2)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "合同单价(元)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "实际工作量(KM2)"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   6360
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
            Picture         =   "frmList.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":1156
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":16F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":1C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":2224
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView List5 
      Height          =   735
      Left            =   360
      TabIndex        =   27
      Top             =   4680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1296
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "收款日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "收款人"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "收款金额(元)"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "收款帐号"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "累计金额(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "录入日期"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.PictureBox PicIncome 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3840
      Width           =   7620
      Begin 合同管理.XButton cmdAddIncome 
         Height          =   300
         Left            =   2520
         TabIndex        =   23
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "增加"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin 合同管理.XButton cmdEditIncome 
         Height          =   300
         Left            =   3360
         TabIndex        =   24
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "编辑"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin 合同管理.XButton cmdDelIncome 
         Height          =   300
         Left            =   4200
         TabIndex        =   25
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "删除"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin VB.Line Line7 
         BorderColor     =   &H00A6A6A6&
         X1              =   334
         X2              =   334
         Y1              =   3
         Y2              =   27
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00A6A6A6&
         X1              =   152
         X2              =   152
         Y1              =   3
         Y2              =   27
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收款列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   26
         Top             =   120
         Width           =   900
      End
      Begin VB.Image ImgIncome 
         Height          =   375
         Left            =   135
         Stretch         =   -1  'True
         Top             =   45
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView List2 
      Height          =   495
      Left            =   600
      TabIndex        =   20
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "项目编号"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "项目名称"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "承包方式"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "承揽人"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "进场人数"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "进场日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "退场日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "工程地点"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "合同总价(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "其它[补贴...](元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "结算价(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "结算日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "预算借支金额(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "付款方式"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "备注"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   16
         Text            =   "录入日期"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.PictureBox PicSub 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Width           =   7620
      Begin 合同管理.XButton cmdDelSub 
         Height          =   300
         Left            =   4200
         TabIndex        =   13
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "删除"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin 合同管理.XButton cmdAddSub 
         Height          =   300
         Left            =   2520
         TabIndex        =   15
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "增加"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin 合同管理.XButton cmdEditSub 
         Height          =   300
         Left            =   3360
         TabIndex        =   16
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "编辑"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin VB.Line Line8 
         BorderColor     =   &H00A6A6A6&
         X1              =   329
         X2              =   329
         Y1              =   2
         Y2              =   26
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00A6A6A6&
         X1              =   147
         X2              =   147
         Y1              =   2
         Y2              =   26
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "子合同列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   1125
      End
      Begin VB.Image ImgSub 
         Height          =   375
         Left            =   60
         Stretch         =   -1  'True
         Top             =   45
         Width           =   375
      End
   End
   Begin VB.PictureBox PicBorrow 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7620
      Begin 合同管理.XButton cmdAddBorrow 
         Height          =   300
         Left            =   2520
         TabIndex        =   10
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "增加"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin 合同管理.XButton cmdEditBorrow 
         Height          =   300
         Left            =   3360
         TabIndex        =   18
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "编辑"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin 合同管理.XButton cmdDelBorrow 
         Height          =   300
         Left            =   4200
         TabIndex        =   19
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "删除"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
         style           =   1
         Enabled         =   0   'False
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
      Begin VB.Line Line5 
         BorderColor     =   &H00A6A6A6&
         X1              =   331
         X2              =   331
         Y1              =   3
         Y2              =   27
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00A6A6A6&
         X1              =   149
         X2              =   149
         Y1              =   3
         Y2              =   27
      End
      Begin VB.Image ImgBorrow 
         Height          =   375
         Left            =   60
         Stretch         =   -1  'True
         Top             =   45
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "借支列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   11
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.ListView List4 
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "借支日期"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "借支人"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "借支金额(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "借支余额(元)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "借支人帐号"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "备注"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "录入日期"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5760
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":27BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmList.frx":3098
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView List1 
      Height          =   540
      Left            =   45
      TabIndex        =   1
      Top             =   480
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   45
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   588
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   8820
      Begin 合同管理.XButton cmdDelMain 
         Height          =   300
         Left            =   6915
         TabIndex        =   7
         Top             =   75
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   "删除"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
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
      Begin 合同管理.XButton cmdAdvSearch 
         Height          =   300
         Left            =   4440
         TabIndex        =   6
         Top             =   75
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         Caption         =   "高级查询"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
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
      Begin 合同管理.FTextBox FTextBox1 
         Height          =   300
         Left            =   1920
         TabIndex        =   5
         ToolTipText     =   "请输入合同名称"
         Top             =   75
         Width           =   1815
         _ExtentX        =   3201
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
         ForeColor       =   12632256
         Text            =   "请输入合同名称"
      End
      Begin 合同管理.XButton cmdSearch 
         Height          =   300
         Left            =   3840
         TabIndex        =   4
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "搜索"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
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
      Begin 合同管理.XButton cmdEditMain 
         Height          =   300
         Left            =   6240
         TabIndex        =   17
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "编辑"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
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
      Begin 合同管理.XButton cmdAddMain 
         Height          =   300
         Left            =   5640
         TabIndex        =   21
         Top             =   75
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         Caption         =   "增加"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
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
      Begin 合同管理.XButton cmdSaveColWidth 
         Height          =   300
         Left            =   7920
         TabIndex        =   29
         Top             =   75
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         Caption         =   "保存列宽"
         BackColor       =   14737632
         ForeColor       =   0
         MouseDownColor  =   255
         MouseOnColor    =   -2147483635
         StyleColor      =   0
         Style3dColor1   =   0
         Style3dColor2   =   0
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
      Begin VB.Line Line3 
         BorderColor     =   &H00A6A6A6&
         X1              =   512
         X2              =   512
         Y1              =   3
         Y2              =   27
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00A6A6A6&
         X1              =   366
         X2              =   366
         Y1              =   3
         Y2              =   27
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合同列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   900
      End
      Begin VB.Image imgIcon 
         Height          =   375
         Left            =   60
         Stretch         =   -1  'True
         Top             =   30
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView List3 
      Height          =   495
      Left            =   2280
      TabIndex        =   28
      Top             =   1680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "工作内容"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "工作量(KM2)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "合同单价(元)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "实际工作量(KM2)"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddBorrow_Click()
    Unload Me
    frmInputBorrow.Show
End Sub

Private Sub cmdAddIncome_Click()
    Unload Me
    frmInputIncome.Show

End Sub

Private Sub cmdAddMain_Click()
    frmMain.cmdLeft_Click 2
End Sub

Private Sub cmdAddSub_Click()
    Unload Me
    frmInputSub.Show
End Sub

Private Sub cmdDelBorrow_Click()
    On Err GoTo errmsg
    Dim i As Integer
    Dim strNO, strSQL As String
    strNO = ""
    
    For i = 1 To List4.ListItems.Count
        If List4.ListItems(i).Selected Then
            strNO = strNO & List4.ListItems(i).Index & "    "
        End If
    Next
    If strNO = "" Then Exit Sub
    
    If MsgBox("确实删除序号为  " & strNO & "的借支记录吗?", vbYesNo + vbExclamation, Me.caption) = vbNo Then Exit Sub
        
        If List4.ListItems.Count > 0 Then
            curList4Index = List4.SelectedItem.Index
        Else
            curList4Index = 0
        End If
        
        DBConnect
        For i = 1 To List4.ListItems.Count
            If List4.ListItems(i).Selected Then
                strSQL = "delete from borrow where id=" & GetID(List4.ListItems(i).Key)
                Conn.Execute strSQL
            End If
        Next
    
        cmdDelBorrow.Enabled = False
        loadBorrowList   '更新借支列表显示
    
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, Me.caption


End Sub

Private Sub cmdDelIncome_Click()
    On Err GoTo errmsg
    Dim i As Integer
    Dim strNO, strSQL As String
    strNO = ""
    
    For i = 1 To List5.ListItems.Count
        If List5.ListItems(i).Selected Then
            strNO = strNO & List5.ListItems(i).Index & "    "
        End If
    Next
    If strNO = "" Then Exit Sub
    
    If MsgBox("确实删除序号为  " & strNO & "的收款记录吗?", vbYesNo + vbExclamation, Me.caption) = vbNo Then Exit Sub
        
        If List5.ListItems.Count > 0 Then
            curList5Index = List5.SelectedItem.Index
        Else
            curList5Index = 0
        End If
        
        DBConnect
        For i = 1 To List5.ListItems.Count
            If List5.ListItems(i).Selected Then
                strSQL = "delete from income where id=" & GetID(List5.ListItems(i).Key)
                Conn.Execute strSQL
            End If
        Next
    
        cmdDelIncome.Enabled = False
        loadIncomeList   '更新收款列表显示
    
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, Me.caption


End Sub

Private Sub cmdDelMain_Click()
    On Err GoTo errmsg
    Dim i As Integer
    Dim strNO, strSQL As String
    strNO = ""
    
    For i = 1 To List1.ListItems.Count
        If List1.ListItems(i).Selected Then
            strNO = strNO & List1.ListItems(i).SubItems(1) & "  "
        End If
    Next
    If strNO = "" Then
        Exit Sub
    End If
    
    If MsgBox("确实删除合同编号为  " & strNO & "的记录吗?", vbYesNo + vbExclamation, Me.caption) = vbNo Then Exit Sub
    
    If List2.ListItems.Count > 0 Then         '存在子合同记录
        If MsgBox("合同  " & strNO & "中包含有" & List2.ListItems.Count & "个子合同，确认删除该合同及其子合同记录吗？", vbYesNo + vbExclamation, Me.caption) = vbNo Then Exit Sub
    End If
        
        If List1.ListItems.Count > 0 Then
            curList1Index = List1.SelectedItem.Index
        Else
            curlist1.Index = 0
        End If
        
        curList2Index = 0
        curList4Index = 0

        
        DBConnect
        
        
        For i = 1 To List2.ListItems.Count                      '删除借支记录
            strSQL = "delete from borrow where zhtid=" & GetID(List2.ListItems(i).Key)
            Conn.Execute strSQL
             strSQL = "delete from subsec where zhtid=" & GetID(List2.ListItems(i).Key)  '删除子合同二记录
            Conn.Execute strSQL
                   
        Next
        
        
        
        strSQL = "delete from sub where zhtid=" & GetID(List1.SelectedItem.Key)   '删除子合同记录
        Conn.Execute strSQL
        
        strSQL = "delete from Income where zhtid=" & GetID(List1.SelectedItem.Key)   '删除收款记录
        Conn.Execute strSQL
        
        strSQL = "delete from mainsec where zhtid=" & GetID(List1.SelectedItem.Key)   '删除主合同二记录
        Conn.Execute strSQL
        
        strSQL = "delete from main where id=" & GetID(List1.SelectedItem.Key)   '删除总合同记录
        Conn.Execute strSQL
        
        Conn.Close
        
        cmdDelMain.Enabled = False
        LoadMainList
        loadSubList
        loadSubSecList
        loadBorrowList
        loadIncomeList
        
    
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, Me.caption
End Sub

Private Sub cmdDelSub_Click()
    On Err GoTo errmsg
    Dim i As Integer
    Dim strNO, strSQL As String
    Dim rs As ADODB.Recordset
    
    strNO = ""
    
    For i = 1 To List2.ListItems.Count
        If List2.ListItems(i).Selected Then
            strNO = strNO & List2.ListItems(i).SubItems(1) & "    "
        End If
    Next
    If strNO = "" Then Exit Sub
    
    If MsgBox("确认删除项目编号为  " & strNO & "的记录吗?", vbYesNo + vbExclamation, Me.caption) = vbNo Then Exit Sub
        
    If List4.ListItems.Count > 0 Then         '存在借支记录
        If MsgBox("项目  " & strNO & "中包含有" & List4.ListItems.Count & "条借支记录，确认删除该项目及其借支记录吗？", vbYesNo + vbExclamation, Me.caption) = vbNo Then Exit Sub
    End If
        
    DBConnect
    If List2.ListItems.Count > 0 Then
        curList2Index = List2.SelectedItem.Index
    Else
        curList2Index = 0
    End If
        
    curList4Index = 0
       
    For i = 1 To List2.ListItems.Count
        If List2.ListItems(i).Selected Then
            strSQL = "delete from borrow where zhtid=" & GetID(List2.ListItems(i).Key)
            Conn.Execute strSQL
            
            strSQL = "delete from subsec where zhtid=" & GetID(List2.ListItems(i).Key)
            Conn.Execute strSQL
            
            strSQL = "delete from sub where id=" & GetID(List2.ListItems(i).Key)
            Conn.Execute strSQL
        End If
    Next
    
    cmdDelSub.Enabled = False
    
    loadSubList
    loadSubSecList
    
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, Me.caption
    
End Sub

Private Sub cmdEditBorrow_Click()
    DataOperateState = "EDIT"
    Unload Me
    frmInputBorrow.Show
End Sub

Private Sub cmdEditIncome_Click()
    DataOperateState = "EDIT"
    Unload Me
    frmInputIncome.Show
End Sub

Private Sub cmdEditMain_Click()
On Error GoTo aaaa
    DataOperateState = "EDIT"
    mainID = GetID(List1.SelectedItem.Key)
    Unload Me
    frmMain.cmdLeft_Click 2
aaaa:

End Sub

Private Sub cmdEditSub_Click()
    DataOperateState = "EDIT"
    Unload Me
    frmInputSub.Show

End Sub

Private Sub cmdSaveColWidth_Click()
    SaveListColWidth   '保存各列表列宽数据
End Sub

Private Sub cmdSearch_Click()
    LoadAllList "htmc", Trim(FTextBox1.Text)  '查询合同名称
End Sub

Private Sub Form_Activate()
    
    SetListColWidth     '设置各列表列宽
    
    LoadAllList      '加载各合同列表
    
End Sub

Private Sub Form_Load()
'On Error GoTo aaaa
    Me.WindowState = vbMaximized    '最大化窗口
    imgIcon.Picture = ImageList2.ListImages(1).Picture
    ImgSub.Picture = ImageList2.ListImages(1).Picture
    ImgBorrow.Picture = ImageList2.ListImages(2).Picture
    ImgIncome.Picture = ImageList2.ListImages(2).Picture
        
    Me.BackColor = color(0)
    List1.BackColor = Me.BackColor
    List2.BackColor = List1.BackColor
    List3.BackColor = List1.BackColor
    List4.BackColor = List1.BackColor
    List5.BackColor = List1.BackColor
    List6.BackColor = List1.BackColor
    
    
    cmdDelMain.Enabled = False
    
    
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
End Sub

Public Sub LoadMainList(Optional strSeachField As String = "", Optional strSeachKey As String = "")
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
    
    If strSeachField = "" And strSeachKey = "" Then
        rs.Open "Select * From main order by lrrq desc,htbh", Conn, 1, 1
    Else
        rs.Open "Select * From main where " & strSeachField & " like '%" & strSeachKey & "%' order by lrrq desc,htbh", Conn, 1, 1
    End If
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
            textcolor = color(1)
        End If
        
        For i = 1 To List1.ColumnHeaders.Count - 2
            List1.ListItems(No).ForeColor = textcolor
            List1.ListItems(No).ListSubItems.Item(i).ForeColor = textcolor
        Next
        
        rs.MoveNext
        
    Loop
    
    cmdDelMain.Enabled = IIf(rs.RecordCount > 0, True, False)
    cmdEditMain.Enabled = cmdDelMain.Enabled
    cmdAddSub.Enabled = IIf(rs.RecordCount > 0, True, False)
    cmdAddIncome.Enabled = IIf(rs.RecordCount > 0, True, False)
    'cmdAddBorrow.Enabled = rs.RecordCount
    
    SetSB 2, "共 " & rs.RecordCount & " 条记录."
    loadMainSecList
    loadSubList
    loadIncomeList
    
    
End Sub
Sub loadMainSecList()
    Dim Item As ListItem
    Dim strSQL, strFormat As String
    Dim rs As ADODB.Recordset
    Dim iNo As Integer

    List6.ListItems.Clear
    
    If List1.ListItems.Count < 1 Then Exit Sub
    
    'strFormat = "0;1;0.000;##,##0.00;0.000;yyyy-mm-dd hh:mm:ss"
    strFormat = ";;0.000;##,##0.00;0.000;yyyy-mm-dd hh:mm:ss"
    strFormat = Replace(strFormat, ";0.000", GetDecLen(bytAfterDec))
    
    mainID = GetID(List1.SelectedItem.Key)
    DBConnect
    Set rs = New ADODB.Recordset
    strSQL = "select * from mainsec where zhtid=" & mainID
    rs.Open strSQL, Conn, 1, 1
    iNo = 0
    arrayFormat = Split(strFormat, ";")
    
    Do While Not rs.EOF
        iNo = iNo + 1
        
        Set Item = List6.ListItems.Add(, , Trim(CStr(rs("fxny"))))
            
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
        
        For i = 1 To List6.ColumnHeaders.Count - 1
            List6.ListItems(iNo).ForeColor = List1.SelectedItem.ForeColor
            List6.ListItems(iNo).ListSubItems.Item(i).ForeColor = List1.SelectedItem.ForeColor
        Next
        
        rs.MoveNext
        
    Loop
    rs.Close
    

End Sub

Sub loadSubList()
    Dim Item As ListItem
    Dim strSQL, strFormat  As String
    Dim rs As ADODB.Recordset
    Dim iNo As Integer
    
    List2.ListItems.Clear


    If List1.ListItems.Count < 1 Then GoTo exitSub
    
    'strFormat = "0;1;2;3;4;5;yyyy-mm-dd;yyyy-mm-dd;8;##,##0.00;##,##0.00;##,##0.00;yyyy-mm-dd;##,##0.00;14;yyyy-mm-dd hh:mm:ss"
    strFormat = ";;;;;;yyyy-mm-dd;yyyy-mm-dd;;##,##0.00;##,##0.00;##,##0.00;yyyy-mm-dd;##,##0.00;;;yyyy-mm-dd hh:mm:ss"
    strFormat = Replace(strFormat, ";0.000", GetDecLen(bytAfterDec))
    arrayFormat = Split(strFormat, ";")
    
    
    mainID = GetID(List1.SelectedItem.Key)
    DBConnect
    Set rs = New ADODB.Recordset
    strSQL = "select * from sub where zhtid=" & mainID & " order by xmbh"
    rs.Open strSQL, Conn, 1, 1
    iNo = 0
    
    
    Do While Not rs.EOF
        iNo = iNo + 1
        
        Set Item = List2.ListItems.Add(, Trim(CStr(rs("id"))) & "k", iNo, , 3)
            
        If Not IsNull(rs("xmbh")) Then Item.SubItems(1) = rs("xmbh")
        If Not IsNull(rs("xmmc")) Then Item.SubItems(2) = rs("xmmc")
        If Not IsNull(rs("cbfs")) Then
            For k = 0 To UBound(strMode)       '获取承包方式
                If rs("cbfs") = strMode(k, 1) Then Item.SubItems(3) = strMode(k, 0)
            Next
        End If
        
        For i = 4 To rs.Fields.Count - 3
            
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
            
            Item.SubItems(i) = Format(temp, arrayFormat(i))
            
        Next
        
        If rs("yjs") Then
            textcolor = color(2)
        Else
            textcolor = color(1)
        End If
        
        For i = 1 To List2.ColumnHeaders.Count - 1
            List2.ListItems(iNo).ForeColor = textcolor
            List2.ListItems(iNo).ListSubItems.Item(i).ForeColor = textcolor
        Next
        
        
        rs.MoveNext
        
    Loop
    
    cmdAddBorrow.Enabled = IIf(rs.RecordCount > 0, True, False)
    cmdEditSub.Enabled = IIf(rs.RecordCount > 0, True, False)
    cmdDelSub.Enabled = IIf(rs.RecordCount > 0, True, False)
    
    
    
exitSub:
    loadBorrowList

End Sub
Sub loadSubSecList()
    Dim Item As ListItem
    Dim strSQL, strFormat As String
    Dim rs As ADODB.Recordset
    Dim iNo As Integer
    
    List3.ListItems.Clear
    
    If List2.ListItems.Count < 1 Then Exit Sub
    
    'strFormat = "0;1;0.000;##,##0.00;0.000;yyyy-mm-dd hh:mm:ss"
    strFormat = ";;0.000;##,##0.00;0.000;yyyy-mm-dd hh:mm:ss"
    strFormat = Replace(strFormat, ";0.000", GetDecLen(bytAfterDec))
    arrayFormat = Split(strFormat, ";")

    
    subID = GetID(List2.SelectedItem.Key)
    DBConnect
    Set rs = New ADODB.Recordset
    strSQL = "select * from subsec where zhtid=" & subID
    rs.Open strSQL, Conn, 1, 1
    iNo = 0
    
    Do While Not rs.EOF
        iNo = iNo + 1
        
        Set Item = List3.ListItems.Add(, , Trim(CStr(rs("gzny"))))
            
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
            
        For i = 1 To List3.ColumnHeaders.Count - 1
            List3.ListItems(iNo).ForeColor = List2.SelectedItem.ForeColor
            List3.ListItems(iNo).ListSubItems.Item(i).ForeColor = List2.SelectedItem.ForeColor
        Next
        
        rs.MoveNext
        
    Loop

End Sub

Sub loadBorrowList()
    Dim Item As ListItem
    Dim strSQL, strFormat As String
    Dim rs As ADODB.Recordset
    Dim iNo As Integer
    Dim dblBorrow As Double      '借支余额
    
    List4.ListItems.Clear
    
    If List2.ListItems.Count < 1 Then Exit Sub
    
    'strFormat = "0;yyyy-mm-dd;2;##,##0.00;##,##0.00;5;6;yyyy-mm-dd hh:mm:ss"
    strFormat = ";yyyy-mm-dd;;##,##0.00;##,##0.00;;;yyyy-mm-dd hh:mm:ss"
    strFormat = Replace(strFormat, ";0.000", GetDecLen(bytAfterDec))
    arrayFormat = Split(strFormat, ";")

    
    subID = GetID(List2.SelectedItem.Key)
    DBConnect
    Set rs = New ADODB.Recordset
    strSQL = "select * from borrow where zhtid=" & subID & " order by jzrq,lrrq"
    rs.Open strSQL, Conn, 1, 1
    iNo = 0
    
    If Trim(List2.SelectedItem.SubItems(13)) <> "" Then
        dblBorrow = CDbl(List2.SelectedItem.SubItems(13))
    Else
        dblBorrow = 0
    End If
    Do While Not rs.EOF
        iNo = iNo + 1
        
        Set Item = List4.ListItems.Add(, Trim(CStr(rs("id"))) & "k", iNo, , 4)
            
        For i = 1 To rs.Fields.Count - 2
            
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
            
            Item.SubItems(i) = Format(temp, arrayFormat(i))
            
        Next
        
        temp = Trim(Item.SubItems(3))
        If temp = "" Then
            temp = 0
        End If
        Item.SubItems(4) = Format(CStr(dblBorrow - temp), arrayFormat(4))
        dblBorrow = Item.SubItems(4)
        
        For i = 1 To List4.ColumnHeaders.Count - 1
            List4.ListItems(iNo).ForeColor = color(1)
            List4.ListItems(iNo).ListSubItems.Item(i).ForeColor = color(1)
        Next
        
        
        rs.MoveNext
        
    Loop
    
    cmdEditBorrow.Enabled = IIf(rs.RecordCount > 0, True, False)
    cmdDelBorrow.Enabled = IIf(rs.RecordCount > 0, True, False)
    
    SetCmdState
    

End Sub
Sub loadIncomeList()
    Dim Item As ListItem
    Dim strSQL, strFormat As String
    Dim rs As ADODB.Recordset
    Dim iNo As Integer
    Dim dblSum As Double '累计金额
    
    List5.ListItems.Clear
    
    If List1.ListItems.Count < 1 Then Exit Sub
    
    'strFormat = "0;yyyy-mm-dd;2;##,##0.00;4;yyyy-mm-dd hh:mm:ss"
    strFormat = ";yyyy-mm-dd;;##,##0.00;;##,##0.00;yyyy-mm-dd hh:mm:ss"  '列表列比数据表多一项累计金额
    strFormat = Replace(strFormat, ";0.000", GetDecLen(bytAfterDec))
    arrayFormat = Split(strFormat, ";")

    
    mainID = GetID(List1.SelectedItem.Key)
    DBConnect
    Set rs = New ADODB.Recordset
    strSQL = "select * from income where zhtid=" & mainID & " order by skrq,lrrq"
    rs.Open strSQL, Conn, 1, 1
    iNo = 0
    dblSum = 0
    
    Do While Not rs.EOF
        iNo = iNo + 1
        
        Set Item = List5.ListItems.Add(, Trim(CStr(rs("id"))) & "k", iNo, , 5)
            
        For i = 1 To rs.Fields.Count - 3
            
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
            
            Item.SubItems(i) = Format(temp, arrayFormat(i))
            
        Next
            
        dblSum = dblSum + Item.SubItems(3)
            
        Item.SubItems(i) = Format(dblSum, arrayFormat(i))
        Item.SubItems(i + 1) = Format(rs("lrrq"), arrayFormat(i + 1)) '录入日期
        
        For i = 1 To List5.ColumnHeaders.Count - 1
            List5.ListItems(iNo).ForeColor = color(1)
            List5.ListItems(iNo).ListSubItems.Item(i).ForeColor = color(1)
        Next
        
        
        rs.MoveNext
        
    Loop
    
    cmdEditIncome.Enabled = IIf(rs.RecordCount > 0, True, False)
    cmdDelIncome.Enabled = IIf(rs.RecordCount > 0, True, False)


End Sub
Private Sub Form_Resize()
On Error Resume Next
    Dim frmWidth, frmHeight As Integer
    Dim intRange As Integer
    intRange = 5
    
    frmWidth = Width / 15 - 16
    frmHeight = Height / 15 - 35
    
    List1.Width = frmWidth * 3 / 4
    List1.Height = frmHeight / 2 - PicTop.Height
    
    MG2.Width = Width / 15 - 16
    MG2.Height = Height / 15 - 30
    PicTop.Width = Width / 15 - 14
    
    List6.Left = List1.Left + List1.Width + intRange
    List6.Top = List1.Top
    List6.Width = frmWidth - List1.Width - intRange
    List6.Height = List1.Height / 2
    
    PicSub.Left = PicTop.Left
    PicSub.Top = List1.Top + List1.Height
    PicSub.Width = PicTop.Width
    
    List2.Left = List1.Left
    List2.Height = frmHeight / 4 - PicSub.Height
    List2.Top = PicSub.Top + PicSub.Height
    List2.Width = frmWidth * 3 / 4
    
    List3.Left = List2.Left + List2.Width + intRange
    List3.Height = List2.Height
    List3.Top = List2.Top
    List3.Width = frmWidth - List2.Width - intRange
    
    PicBorrow.Left = PicTop.Left
    PicBorrow.Top = List2.Top + List2.Height
    PicBorrow.Width = PicTop.Width / 2
    
    List4.Left = List2.Left
    List4.Height = frmHeight / 4 - PicBorrow.Height
    List4.Top = PicBorrow.Top + PicBorrow.Height
    List4.Width = frmWidth / 2
    
    PicIncome.Left = PicBorrow.Left + PicBorrow.Width
    PicIncome.Top = PicBorrow.Top
    PicIncome.Width = PicBorrow.Width
    
    List5.Left = List4.Left + List4.Width + intRange
    List5.Height = List4.Height
    List5.Top = List4.Top
    List5.Width = List4.Width - intRange
    
    
    
    
    Cls
    Line (2, 2)-(Width / 15 - 12, Height / 15 - 29), 10921638, B
End Sub


Private Sub freItem_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If List1.ListItems.Count > 0 Then
        mainID = GetID(List1.SelectedItem.Key)
        curList1Index = List1.SelectedItem.Index
    End If
    
    If List2.ListItems.Count > 0 Then
        subID = GetID(List2.SelectedItem.Key)
        curList2Index = List2.SelectedItem.Index
    End If
    
    If List4.ListItems.Count > 0 Then
        borrowID = GetID(List4.SelectedItem.Key)
        dblBalace = List4.SelectedItem.SubItems(4)   '获取借支余额
        curList4Index = List4.SelectedItem.Index
    End If
    
    If List5.ListItems.Count > 0 Then
        incomeID = GetID(List5.SelectedItem.Key)
        curList5Index = List5.SelectedItem.Index
    End If
    
    SetCmdEnable (False)
    
    SetSB 2, ""
    

End Sub

Private Sub FTextBox1_Click()
    If FTextBox1.Text = "请输入合同名称" Then
        FTextBox1.Text = ""
        FTextBox1.ForeColor = RGB(50, 50, 50)
    End If
End Sub

Private Sub List1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If List1.SortOrder = lvwDescending Then
        List1.SortOrder = lvwAscending
    Else
        List1.SortOrder = lvwDescending
    End If
    
    List1.SortKey = ColumnHeader.Index - 1
    List1.Sorted = True
End Sub

Private Sub List1_DblClick()
    cmdEditMain_Click
End Sub
Private Sub List1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdDelMain.Enabled = True
    loadSubList
    loadSubSecList
    
    frmMain.cmdLeft(5).Enabled = List2.ListItems.Count
    frmMain.tbLeft(5).Enabled = List2.ListItems.Count
    frmMain.mnuLeft(5).Enabled = List2.ListItems.Count
    
    loadMainSecList
    loadIncomeList
    
    SetCmdState
    
  
End Sub
Private Sub List2_DblClick()
    cmdEditSub_Click
End Sub

Private Sub List2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdDelSub.Enabled = True
    loadSubSecList
    loadBorrowList
End Sub
Private Sub List4_DblClick()
    cmdEditBorrow_Click
End Sub

Private Sub List4_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdDelBorrow.Enabled = True
End Sub

Private Sub List41_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub List5_DblClick()
    cmdEditIncome_Click
End Sub
Sub SetCmdState()
        
    If List1.ListItems.Count > 0 And List2.ListItems.Count > 0 Then
        frmMain.cmdLeft(3).Enabled = True '结算单
        frmMain.mnuExItem.Enabled = True  '导出项目资料
    Else
        frmMain.cmdLeft(3).Enabled = False
        frmMain.mnuExItem.Enabled = False
    End If
    
    If List1.ListItems.Count > 0 And List2.ListItems.Count > 0 And List3.ListItems.Count > 0 Then
        frmMain.cmdLeft(4).Enabled = True '确认单
    Else
        frmMain.cmdLeft(4).Enabled = False
    End If
    
    frmMain.tbLeft(3).Enabled = frmMain.cmdLeft(3).Enabled
    frmMain.tbLeft(4).Enabled = frmMain.cmdLeft(4).Enabled
    frmMain.mnuLeft(3).Enabled = frmMain.cmdLeft(3).Enabled
    frmMain.mnuLeft(4).Enabled = frmMain.cmdLeft(4).Enabled
    
    If List2.ListItems.Count > 0 And List3.ListItems.Count > 0 And List4.ListItems.Count > 0 Then
        frmMain.cmdLeft(5).Enabled = True                            '借支单
    Else
        frmMain.cmdLeft(5).Enabled = False
    End If
    
    frmMain.tbLeft(5).Enabled = frmMain.cmdLeft(5).Enabled
    frmMain.mnuLeft(5).Enabled = frmMain.cmdLeft(5).Enabled
    
    
End Sub
Sub SetListColWidth()
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim col As Byte    '列数
    
    Set rs = New ADODB.Recordset
    
    DBConnect
    
    strSQL = "select * from listpara where listname='list1' order by col"
    rs.Open strSQL, Conn, 1, 1
    col = List1.ColumnHeaders.Count
    n = 0
    Do While Not rs.EOF
        n = n + 1
        If Not IsNull(rs("width")) And n <= col Then
            List1.ColumnHeaders.Item(n).Width = rs("width")
        End If
            
        rs.MoveNext
    Loop
    rs.Close
    List1.Refresh
    
    strSQL = "select * from listpara where listname='list2' order by col"
    rs.Open strSQL, Conn, 1, 1
    col = List2.ColumnHeaders.Count
    
    n = 0
    Do While Not rs.EOF
        n = n + 1
        If Not IsNull(rs("width")) And n <= col Then
            List2.ColumnHeaders.Item(n).Width = rs("width")
        End If
            
        rs.MoveNext
    Loop
    rs.Close
    List2.Refresh
    
    strSQL = "select * from listpara where listname='list3' order by col"
    rs.Open strSQL, Conn, 1, 1
    col = List3.ColumnHeaders.Count
    
    n = 0
    Do While Not rs.EOF
        n = n + 1
        If Not IsNull(rs("width")) And n <= col Then
            List3.ColumnHeaders.Item(n).Width = rs("width")
        End If
            
        rs.MoveNext
    Loop
    rs.Close
    List3.Refresh
    
    strSQL = "select * from listpara where listname='list4' order by col"
    rs.Open strSQL, Conn, 1, 1
    col = List4.ColumnHeaders.Count
    
    n = 0
    Do While Not rs.EOF
        n = n + 1
        If Not IsNull(rs("width")) And n <= col Then
            List4.ColumnHeaders.Item(n).Width = rs("width")
        End If
            
        rs.MoveNext
    Loop
    rs.Close
    List4.Refresh
    
    strSQL = "select * from listpara where listname='list5' order by col"
    rs.Open strSQL, Conn, 1, 1
    col = List5.ColumnHeaders.Count
    
    n = 0
    Do While Not rs.EOF
        n = n + 1
        If Not IsNull(rs("width")) And n <= col Then
            List5.ColumnHeaders.Item(n).Width = rs("width")
        End If
            
        rs.MoveNext
    Loop
    rs.Close
    List5.Refresh
    
    strSQL = "select * from listpara where listname='list6' order by col"
    rs.Open strSQL, Conn, 1, 1
    col = List6.ColumnHeaders.Count
    
    n = 0
    Do While Not rs.EOF
        n = n + 1
        If Not IsNull(rs("width")) And n <= col Then
            List6.ColumnHeaders.Item(n).Width = rs("width")
        End If
            
        rs.MoveNext
    Loop
    rs.Close
    List6.Refresh
    Conn.Close

End Sub
Sub SaveListColWidth()
    Dim strSQL As String
    Dim i As Byte
    
    DBConnect
    strSQL = "delete from listpara"
    Conn.Execute strSQL
    For i = 1 To List1.ColumnHeaders.Count
        strSQL = "insert into listpara(listname,col,width) values('list1'," & i & "," & List1.ColumnHeaders.Item(i).Width & ")"
    Conn.Execute strSQL
    Next
    
    For i = 1 To List2.ColumnHeaders.Count
        strSQL = "insert into listpara(listname,col,width) values('list2'," & i & "," & List2.ColumnHeaders.Item(i).Width & ")"
    Conn.Execute strSQL
    Next
    
    For i = 1 To List3.ColumnHeaders.Count
        strSQL = "insert into listpara(listname,col,width) values('list3'," & i & "," & List3.ColumnHeaders.Item(i).Width & ")"
    Conn.Execute strSQL
    Next
    
    For i = 1 To List4.ColumnHeaders.Count
        strSQL = "insert into listpara(listname,col,width) values('list4'," & i & "," & List4.ColumnHeaders.Item(i).Width & ")"
    Conn.Execute strSQL
    Next
    
    For i = 1 To List5.ColumnHeaders.Count
        strSQL = "insert into listpara(listname,col,width) values('list5'," & i & "," & List5.ColumnHeaders.Item(i).Width & ")"
    Conn.Execute strSQL
    Next
    
    For i = 1 To List6.ColumnHeaders.Count
        strSQL = "insert into listpara(listname,col,width) values('list6'," & i & "," & List6.ColumnHeaders.Item(i).Width & ")"
    Conn.Execute strSQL
    Next
    
    Conn.Close

End Sub
Sub LoadList()
    Form_Activate
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

Sub SetCmdEnable(boolEnable As Boolean)
    Dim i As Integer
    For i = 3 To 7
        If i <> 6 Then  '用户设置功能不进行设置
            frmMain.cmdLeft(i).Enabled = boolEnable
            frmMain.tbLeft(i).Enabled = boolEnable
            frmMain.mnuLeft(i).Enabled = boolEnable
        End If
    Next
    
    frmMain.mnuExItem.Enabled = boolEnable

End Sub

Sub LoadAllList(Optional strSeachField As String = "", Optional strSeachKey As String = "")
    
    LoadMainList strSeachField, strSeachKey   '加载合同列表
    If curList1Index > 0 And curList1Index <= List1.ListItems.Count Then
        List1.ListItems(curList1Index).Selected = True
    End If
    
    loadMainSecList  '加载收款列表
    If curList6Index > 0 And curList6Index <= List6.ListItems.Count Then
        List6.ListItems(curList6Index).Selected = True
    End If
    
    loadIncomeList  '加载收款列表
    If curList5Index > 0 And curList5Index <= List5.ListItems.Count Then
        List5.ListItems(curList5Index).Selected = True
    End If
    
    
    loadSubList     '加载子合同列表
    If curList2Index > 0 And curList2Index <= List2.ListItems.Count Then
        List2.ListItems(curList2Index).Selected = True
    End If
    
    loadSubSecList  '加载子合同二列表
    If curList3Index > 0 And curList3Index <= List3.ListItems.Count Then
        List3.ListItems(curList3Index).Selected = True
    End If
    
    
    loadBorrowList   '加载借支列表
    If curList4Index > 0 And curList4Index <= List4.ListItems.Count Then
        List4.ListItems(curList4Index).Selected = True
    End If
    
    SetCmdEnable True
    
    SetCmdState

End Sub
