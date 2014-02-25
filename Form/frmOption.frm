VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOption 
   AutoRedraw      =   -1  'True
   Caption         =   "选项设置"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   ControlBox      =   0   'False
   Icon            =   "frmOption.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   602
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   600
      Top             =   6930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":173E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOption.frx":1CD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   6615
      Begin 合同管理.XPButton cmdExitOption 
         Height          =   345
         Index           =   1
         Left            =   4440
         TabIndex        =   35
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "返回(&Q)"
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
      Begin VB.Frame freItem 
         Height          =   2535
         Index           =   1
         Left            =   1080
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   4380
         Begin 合同管理.FTextBox txtName 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   9
            Top             =   840
            Width           =   2835
            _ExtentX        =   5001
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
            AutoSelAll      =   -1  'True
         End
         Begin 合同管理.FTextBox txtID 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   8
            Top             =   360
            Width           =   2835
            _ExtentX        =   5001
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
            Enabled         =   0   'False
            AutoSelAll      =   -1  'True
            isNumber        =   -1  'True
            MaxLength       =   2
         End
         Begin 合同管理.FTextBox txtDesc 
            Height          =   300
            Index           =   1
            Left            =   1200
            TabIndex        =   10
            Top             =   1320
            Width           =   2835
            _ExtentX        =   5001
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
            AutoSelAll      =   -1  'True
         End
         Begin 合同管理.XPButton cmdExit 
            Height          =   345
            Index           =   1
            Left            =   2940
            TabIndex        =   17
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "取消"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin 合同管理.XPButton cmdOK 
            Height          =   345
            Index           =   1
            Left            =   1740
            TabIndex        =   18
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "添加"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型说明"
            Height          =   180
            Left            =   360
            TabIndex        =   21
            Top             =   1395
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型编号"
            Height          =   180
            Left            =   360
            TabIndex        =   20
            Top             =   435
            Width           =   720
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型名称"
            Height          =   180
            Left            =   360
            TabIndex        =   19
            Top             =   915
            Width           =   720
         End
      End
      Begin 合同管理.XPButton cmdDel 
         Height          =   345
         Index           =   1
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "删除(&D)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin 合同管理.XPButton cmdEdit 
         Height          =   345
         Index           =   1
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "修改(&E)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin 合同管理.XPButton cmdAdd 
         Height          =   345
         Index           =   1
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "添加(&A)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3615
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6376
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
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "类型编号"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "类型名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "备注"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   2
      Left            =   435
      TabIndex        =   6
      Top             =   1035
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame freItem 
         Height          =   2535
         Index           =   2
         Left            =   1080
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   4380
         Begin 合同管理.FTextBox txtName 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   27
            Top             =   840
            Width           =   2835
            _ExtentX        =   5001
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
            AutoSelAll      =   -1  'True
         End
         Begin 合同管理.FTextBox txtID 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   28
            Top             =   360
            Width           =   2835
            _ExtentX        =   5001
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
            AutoSelAll      =   -1  'True
            isNumber        =   -1  'True
            MaxLength       =   2
         End
         Begin 合同管理.FTextBox txtDesc 
            Height          =   300
            Index           =   2
            Left            =   1200
            TabIndex        =   29
            Top             =   1320
            Width           =   2835
            _ExtentX        =   5001
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
            AutoSelAll      =   -1  'True
         End
         Begin 合同管理.XPButton cmdExit 
            Height          =   345
            Index           =   2
            Left            =   2940
            TabIndex        =   30
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "取消"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin 合同管理.XPButton cmdOK 
            Height          =   345
            Index           =   2
            Left            =   1740
            TabIndex        =   31
            Top             =   1950
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Caption         =   "添加"
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mode            =   0
            Value           =   0   'False
            cBack           =   -2147483633
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型名称"
            Height          =   180
            Left            =   360
            TabIndex        =   34
            Top             =   915
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型编号"
            Height          =   180
            Left            =   360
            TabIndex        =   33
            Top             =   435
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "类型说明"
            Height          =   180
            Left            =   360
            TabIndex        =   32
            Top             =   1395
            Width           =   720
         End
      End
      Begin MSComctlLib.ListView List1 
         Height          =   3615
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6376
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
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "类型编号"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "类型名称"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "备注"
            Object.Width           =   3528
         EndProperty
      End
      Begin 合同管理.XPButton cmdDel 
         Height          =   345
         Index           =   2
         Left            =   3240
         TabIndex        =   23
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "删除(&D)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin 合同管理.XPButton cmdEdit 
         Height          =   345
         Index           =   2
         Left            =   2040
         TabIndex        =   24
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "修改(&E)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin 合同管理.XPButton cmdAdd 
         Height          =   345
         Index           =   2
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "添加(&A)"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin 合同管理.XPButton cmdExitOption 
         Height          =   345
         Index           =   2
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "返回(&Q)"
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
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   3
      Left            =   1605
      TabIndex        =   1
      Top             =   2430
      Visible         =   0   'False
      Width           =   6615
      Begin 合同管理.FCombo CboDec 
         Height          =   300
         Left            =   3240
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
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
         EnabledText     =   0   'False
         ListIndex       =   -1
      End
      Begin 合同管理.XPButton cmdExitOption 
         Height          =   345
         Index           =   3
         Left            =   3480
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "返回(&Q)"
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
      Begin 合同管理.XPButton cmdSaveCon 
         Height          =   345
         Left            =   1920
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "保存(&S)"
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
      Begin VB.Label lblInfo 
         Caption         =   "参数保存完成!"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "工作量小数位数"
         Height          =   225
         Left            =   1800
         TabIndex        =   3
         Top             =   1245
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Index           =   4
      Left            =   3000
      TabIndex        =   37
      Top             =   4320
      Visible         =   0   'False
      Width           =   6615
      Begin 合同管理.XPButton cmdSet 
         Height          =   270
         Index           =   2
         Left            =   4710
         TabIndex        =   49
         Top             =   2550
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
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
      Begin 合同管理.XPButton cmdSet 
         Height          =   270
         Index           =   0
         Left            =   4710
         TabIndex        =   48
         Top             =   1320
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
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
         cBhover         =   -2147483635
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16053492
      End
      Begin 合同管理.FTextBox txtColor 
         Height          =   300
         Index           =   2
         Left            =   3105
         TabIndex        =   46
         Top             =   2535
         Width           =   1935
         _ExtentX        =   3413
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
         MaxLength       =   9
         afterdecimal    =   0
      End
      Begin 合同管理.FTextBox txtColor 
         Height          =   300
         Index           =   0
         Left            =   3075
         TabIndex        =   43
         Top             =   1305
         Width           =   1935
         _ExtentX        =   3413
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
         MaxLength       =   9
         afterdecimal    =   0
      End
      Begin 合同管理.XPButton cmdExitOption 
         Height          =   345
         Index           =   0
         Left            =   4380
         TabIndex        =   38
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "返回(&Q)"
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
         Height          =   345
         Left            =   1575
         TabIndex        =   39
         Top             =   3390
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "保存(&S)"
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
      Begin 合同管理.XPButton XPButton2 
         Height          =   345
         Left            =   2970
         TabIndex        =   47
         Top             =   3375
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
         Caption         =   "缺省值(&D)"
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
      Begin 合同管理.XPButton cmdSet 
         Height          =   270
         Index           =   1
         Left            =   4710
         TabIndex        =   50
         Top             =   1920
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   476
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
      Begin 合同管理.FTextBox txtColor 
         Height          =   300
         Index           =   1
         Left            =   3105
         TabIndex        =   44
         Top             =   1905
         Width           =   1935
         _ExtentX        =   3413
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
         MaxLength       =   9
         afterdecimal    =   0
      End
      Begin VB.Label Label12 
         Caption         =   "列表背景颜色"
         Height          =   270
         Left            =   1920
         TabIndex        =   45
         Top             =   1365
         Width           =   1155
      End
      Begin VB.Label Label11 
         Caption         =   "已结算文本颜色"
         Height          =   270
         Left            =   1755
         TabIndex        =   42
         Top             =   2610
         Width           =   1320
      End
      Begin VB.Label Label10 
         Caption         =   "列表文本颜色"
         Height          =   285
         Left            =   1905
         TabIndex        =   41
         Top             =   1950
         Width           =   1170
      End
      Begin VB.Label Label9 
         Caption         =   "颜色设置保存完成!"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2160
         TabIndex        =   40
         Top             =   420
         Visible         =   0   'False
         Width           =   2730
      End
   End
   Begin MSComctlLib.TabStrip tabOption 
      Height          =   735
      Left            =   5100
      TabIndex        =   0
      Top             =   6870
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      TabStyle        =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "合同类型"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "承包方式"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "工作量"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "文字颜色"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intCurFrame As Integer     '当前显示的frame
Private Sub CboDec_Click()
    lblInfo.Visible = False
End Sub

Private Sub cmdSave_Click()
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    cmdOK(Index).caption = "添加"
    List1(Index).Visible = False
    freItem(Index).Visible = True
    txtID(Index).Enabled = True
    txtID(Index).SetFocus
End Sub

Private Sub cmdDel_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    Select Case Index
        Case 1
            rs.Open "select top 1 id from main where htlx=" & List1(Index).SelectedItem.SubItems(1), Conn, 1, 1
        Case 2
            rs.Open "select top 1 id from sub where cbfs=" & List1(Index).SelectedItem.SubItems(1), Conn, 1, 1
    End Select
        
    If Not rs.EOF Then
        MsgBox "已经使用了的类型编号不能删除！", vbExclamation, "选项设置"
        rs.Close
        Exit Sub
    End If
    rs.Close
    
    If MsgBox("确实删除类型名称为 [" & List1(Index).SelectedItem.SubItems(2) & "] 的项目吗？", vbExclamation + vbYesNo, "选项设置") = vbNo Then Exit Sub
    
    Conn.Execute "delete from ItemInfo where id=" & GetID(List1(Index).SelectedItem.Key)
    Conn.Close
    
    loadItemData Index
    
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    cmdOK(Index).caption = "修改"
    txtID(Index).Text = List1(Index).SelectedItem.SubItems(1)
    txtName(Index).Text = List1(Index).SelectedItem.SubItems(2)
    txtDesc(Index).Text = List1(Index).SelectedItem.SubItems(3)
    txtID(Index).Enabled = False
    
    List1(Index).Visible = False
    freItem(Index).Visible = True
    txtID(Index).Enabled = False
    txtID(Index).SetFocus
    
End Sub

Private Sub cmdExit_Click(Index As Integer)
    freItem(Index).Visible = False
    List1(Index).Visible = True
End Sub

Private Sub cmdExitOption_Click(Index As Integer)
    GetItemInfo
    
    frmMain.cmdLeft_Click 1  '返回合同列表窗口
End Sub

Private Sub cmdOK_Click(Index As Integer)
    On Error GoTo errmsg
    Dim rs As ADODB.Recordset
    
    If Trim(txtID(Index).Text) = "" Or Trim(txtName(Index).Text) = "" Then
        MsgBox "类型编号或类型名称未填写！", vbExclamation, "选项设置"
        txtID(Index).SetFocus
        Exit Sub
    End If
    
    DBConnect
    
    If cmdOK(Index).caption = "添加" Then
        Set rs = New ADODB.Recordset
        rs.Open "select * from ItemInfo where ItemID=" & txtID(Index).Text & " and ItemType=" & intCurFrame, Conn, 1, 1
        If Not rs.EOF Then
            MsgBox "类型编号：" & txtID(Index).Text & " 存在，请使用其他编号。"
            txtID(Index).SetFocus
            Exit Sub
        Else
            Conn.Execute "insert into ItemInfo(ItemType,ItemID,ItemName,ItemDesc) values(" & _
                        intCurFrame & "," & txtID(Index).Text & ",'" & Trim(txtName(Index).Text) & "','" & IIf(Trim(txtDesc(Index).Text) <> "", Trim(txtDesc(Index).Text), "") & "')"
        End If
        
        rs.Close
        
    Else
        Conn.Execute "update ItemInfo set ItemID=" & txtID(Index).Text & "," & _
                                  "ItemName='" & Trim(txtName(Index).Text) & "'," & _
                                  "ItemDesc=" & IIf(Trim(txtDesc(Index).Text) <> "", "'" & Trim(txtDesc(Index).Text) & "'", "NULL") & " " & _
                                  "where id=" & GetID(List1(Index).SelectedItem.Key)
    
    End If
    
    Conn.Close
    
    txtID(Index).Text = ""
    txtName(Index).Text = ""
    txtDesc(Index).Text = ""
    
    loadItemData Index
    
    List1(Index).Visible = True
    freItem(Index).Visible = False
    
    Exit Sub
errmsg:
    MsgBox Err.Description, vbCritical, "选项设置"
    
End Sub

Private Sub cmdSaveCon_Click()
    On Error GoTo errmsg
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    strSQL = "select * from ItemInfo where ItemType=3"
    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then
        strSQL = "insert into ItemInfo(ItemType,ItemValue) values(3," & CboDec.ListIndex & ")"
    Else
        strSQL = "update ItemInfo set ItemValue=" & CboDec.ListIndex & " where ItemType=3"
    End If
    
    rs.Close
    Conn.Execute strSQL
    Conn.Close
    
    lblInfo.Visible = True
    bytAfterDec = CboDec.ListIndex
    Exit Sub
    
errmsg:
    MsgBox Err.Description, vbCritical, "选项设置"

End Sub

Private Sub cmdSet_Click(Index As Integer)
    On Error GoTo errmsg
    Label9.Visible = False
    
    ComDlg.CancelError = True
    ComDlg.ShowColor
    
    txtColor(Index).Text = ComDlg.color
    cmdSet(Index).BackColor = ComDlg.color
    
    Exit Sub
    
errmsg:
    
End Sub

Private Sub Form_Load()
    Width = 6735
    Height = 5600

    For i = 0 To 10
        CboDec.AddItem i, i
    Next
    
    intCurFrame = 1
    loadItemData (1)
    
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'PicTop.Width = Width / 15 - 16
    Cls
    Line (2, 2)-(Width / 15 - 14, Height / 15 - 29), 10921638, B
    tabOption.Width = Width / 15
    tabOption.Top = 0
    tabOption.Left = 0
    tabOption.Height = Height / 15
    
    For i = 1 To 4
        Frame1(i).Top = tabOption.ClientTop
        Frame1(i).Left = tabOption.Left
        Frame1(i).Height = tabOption.Height
        Frame1(i).Width = tabOption.Width
    Next
    
    List1(1).ColumnHeaders.Item(4).Width = List1(1).Width - List1(1).ColumnHeaders.Item(1).Width - List1(1).ColumnHeaders.Item(2).Width - List1(1).ColumnHeaders.Item(3).Width - 90
    List1(2).ColumnHeaders.Item(4).Width = List1(2).Width - List1(2).ColumnHeaders.Item(1).Width - List1(2).ColumnHeaders.Item(2).Width - List1(1).ColumnHeaders.Item(3).Width - 90
    

End Sub

Private Sub tabOption_Click()
    If tabOption.SelectedItem.Index = intCurFrame Then Exit Sub
    Frame1(tabOption.SelectedItem.Index).Visible = True
    Frame1(intCurFrame).Visible = False
    intCurFrame = tabOption.SelectedItem.Index
    loadItemData tabOption.SelectedItem.Index

    
End Sub
Sub loadItemData(intTabIndex As Integer)
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim Item As ListItem
    Dim AfterDec As Integer
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    strSQL = "select * from ItemInfo where ItemType=" & intTabIndex & " order by ItemID"
    
    Select Case intTabIndex
        Case 1, 2        '1-合同类型,2-承包方式
            rs.Open strSQL, Conn, 1, 1
            List1(intTabIndex).ListItems.Clear
            Do While Not rs.EOF
                iNo = iNo + 1
                Set Item = List1(intTabIndex).ListItems.Add(, rs("id") & "k", , , intTabIndex)
                For i = 1 To rs.Fields.Count - 3
                    If Not IsNull(rs.Fields(i).value) Then
                        Item.SubItems(i) = rs.Fields(i).value
                    End If
                Next
                rs.MoveNext
            Loop
            
            If List1(intTabIndex).ListItems.Count > 0 Then
                cmdEdit(intTabIndex).Enabled = True
                cmdDel(intTabIndex).Enabled = True
            Else
                cmdEdit(intTabIndex).Enabled = False
                cmdDel(intTabIndex).Enabled = False
            
            End If
            
        Case 3   '工作量小数位数
            rs.Open strSQL, Conn, 1, 1
            AfterDec = 3
            If Not rs.EOF Then
                If Not IsNull(rs("ItemValue")) Then
                    AfterDec = rs("ItemValue")
                End If
            End If
            
            CboDec.ListIndex = AfterDec
            
        Case 4
            For i = 0 To 2
                strSQL = "select * from ItemInfo where ItemType=4 and ItemID=" & i
                rs.Open strSQL, Conn, 1, 1
                If rs.EOF Then
                    txtColor(i).Text = color(i)
                Else
                    txtColor(i).Text = rs("ItemValue")
                End If
                
                cmdSet(i).BackColor = txtColor(i).Text
                
                rs.Close
            Next
            
            

        End Select
            
    If rs.State <> 0 Then rs.Close
    Conn.Close


End Sub

Private Sub txtColor_Change(Index As Integer)
    If Trim(txtColor(Index).Text) = "" Then Exit Sub
    If CLng(Trim(txtColor(Index).Text)) > 16777215 Then txtColor(Index).Text = 16777215    '大于&HFFFFFF
    If CLng(Trim(txtColor(Index).Text)) < 0 Then txtColor(Index).Text = 0
    
    cmdSet(Index).BackColor = txtColor(Index).Text
    
End Sub

Private Sub txtColor_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Label9.Visible = False
End Sub

Private Sub txtDesc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
         cmdAdd(Index).SetFocus
    End If

End Sub

Private Sub txtID_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtName(Index).SetFocus
    End If
End Sub

Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtDesc(Index).SetFocus
    End If
    
End Sub

Private Sub XPButton1_Click()
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    
    For i = 0 To 2   '0-列表背景色，1-列表文本色，2-已结算文本色
        strSQL = "select id from ItemInfo where ItemType=4 and ItemID=" & i
        rs.Open strSQL, Conn, 1, 1
        If rs.EOF Then
            strSQL = "insert into ItemInfo(ItemType,ItemID,ItemValue) values(4," & i & "," & txtColor(i).Text & ")"
        Else
            strSQL = "update ItemInfo set ItemValue=" & txtColor(i).Text & " " & "where ItemType=4 and ItemID=" & i
        End If
        rs.Close
        Conn.Execute strSQL
        
        color(i) = txtColor(i).Text
        
    Next
    
    Conn.Close
    Label9.Visible = True

End Sub

Private Sub XPButton2_Click()
    txtColor(0).Text = 16777215
    txtColor(1).Text = 0
    txtColor(2).Text = 32768
End Sub
