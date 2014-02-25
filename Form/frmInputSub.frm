VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputSub 
   Caption         =   "数据修改"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   9480
   Begin VB.CheckBox chkFinish 
      Height          =   255
      Left            =   6495
      TabIndex        =   16
      Top             =   3615
      Width           =   375
   End
   Begin VB.PictureBox PicTop 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   644
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   9660
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "子合同信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3600
         TabIndex        =   122
         Top             =   60
         Width           =   2055
      End
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
         TabIndex        =   19
         Top             =   120
         Width           =   900
      End
   End
   Begin 合同管理.XPButton cmdExit 
      Height          =   375
      Left            =   5490
      TabIndex        =   18
      Top             =   8925
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
      Left            =   2970
      TabIndex        =   17
      Top             =   8925
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
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   300
      Index           =   12
      Left            =   8130
      TabIndex        =   21
      Top             =   2355
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   66912257
      CurrentDate     =   39889
   End
   Begin 合同管理.XPButton cmdCLS 
      Height          =   330
      Index           =   6
      Left            =   3495
      TabIndex        =   22
      Top             =   2850
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
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   300
      Index           =   7
      Left            =   3255
      TabIndex        =   23
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   66912257
      CurrentDate     =   39889
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   300
      Index           =   6
      Left            =   3255
      TabIndex        =   24
      Top             =   2880
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   66912257
      CurrentDate     =   39889
   End
   Begin 合同管理.XPButton cmdCLS 
      Height          =   330
      Index           =   7
      Left            =   3495
      TabIndex        =   25
      Top             =   3210
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
   Begin 合同管理.XPButton cmdCLS 
      Height          =   330
      Index           =   12
      Left            =   8370
      TabIndex        =   26
      Top             =   2325
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
   Begin 合同管理.FTextBox txtNo 
      Height          =   300
      Left            =   1230
      TabIndex        =   27
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
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
   End
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   1
      Left            =   1230
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtName 
      Height          =   300
      Left            =   6090
      TabIndex        =   28
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
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
   End
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   9
      Left            =   6090
      TabIndex        =   10
      Top             =   1140
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      BackColor       =   16777215
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
      MaxLength       =   18
   End
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   6
      Left            =   1230
      TabIndex        =   5
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   7
      Left            =   1230
      TabIndex        =   6
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   12
      Left            =   6090
      TabIndex        =   13
      Top             =   2355
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   10
      Left            =   6090
      TabIndex        =   11
      Top             =   1545
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FCombo cobCBFS 
      Height          =   300
      Left            =   1230
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   8
      Left            =   1230
      TabIndex        =   7
      Top             =   3600
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   4
      Left            =   1230
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   5
      Left            =   1230
      TabIndex        =   4
      Top             =   2520
      Width           =   2895
      _ExtentX        =   5106
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
      MaxLength       =   5
      afterdecimal    =   0
   End
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   11
      Left            =   6090
      TabIndex        =   12
      Top             =   1950
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   529
      BackColor       =   16777215
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   13
      Left            =   6090
      TabIndex        =   14
      Top             =   2790
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   16
      Left            =   6090
      TabIndex        =   15
      Top             =   3210
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   2
      Left            =   1230
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   14
      Left            =   1230
      TabIndex        =   8
      Top             =   3960
      Width           =   7755
      _ExtentX        =   13679
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
   Begin 合同管理.XPButton cmdAddRow 
      Height          =   315
      Index           =   1
      Left            =   450
      TabIndex        =   46
      Top             =   5160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "+"
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   2
      Left            =   3810
      TabIndex        =   48
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   3
      Left            =   5370
      TabIndex        =   49
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   1
      Left            =   810
      TabIndex        =   47
      Top             =   5160
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   4
      Left            =   6810
      TabIndex        =   50
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   6
      Left            =   3810
      TabIndex        =   52
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   7
      Left            =   5370
      TabIndex        =   53
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   5
      Left            =   810
      TabIndex        =   51
      Top             =   5520
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   8
      Left            =   6810
      TabIndex        =   54
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   10
      Left            =   3810
      TabIndex        =   56
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   11
      Left            =   5370
      TabIndex        =   57
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   9
      Left            =   810
      TabIndex        =   55
      Top             =   5880
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   12
      Left            =   6810
      TabIndex        =   58
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   14
      Left            =   3810
      TabIndex        =   60
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   15
      Left            =   5370
      TabIndex        =   61
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   13
      Left            =   810
      TabIndex        =   59
      Top             =   6240
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   16
      Left            =   6810
      TabIndex        =   62
      Top             =   6240
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   18
      Left            =   3810
      TabIndex        =   64
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   19
      Left            =   5370
      TabIndex        =   65
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   17
      Left            =   810
      TabIndex        =   63
      Top             =   6600
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   20
      Left            =   6810
      TabIndex        =   66
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   22
      Left            =   3810
      TabIndex        =   68
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   23
      Left            =   5370
      TabIndex        =   69
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   21
      Left            =   810
      TabIndex        =   67
      Top             =   6960
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   24
      Left            =   6810
      TabIndex        =   70
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   26
      Left            =   3810
      TabIndex        =   72
      Top             =   7320
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   27
      Left            =   5370
      TabIndex        =   73
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   25
      Left            =   810
      TabIndex        =   71
      Top             =   7320
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   28
      Left            =   6810
      TabIndex        =   74
      Top             =   7320
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   30
      Left            =   3810
      TabIndex        =   76
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   31
      Left            =   5370
      TabIndex        =   77
      Top             =   7680
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   29
      Left            =   810
      TabIndex        =   75
      Top             =   7680
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   32
      Left            =   6810
      TabIndex        =   78
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   34
      Left            =   3810
      TabIndex        =   80
      Top             =   8040
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   35
      Left            =   5370
      TabIndex        =   81
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   33
      Left            =   810
      TabIndex        =   79
      Top             =   8040
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   36
      Left            =   6810
      TabIndex        =   82
      Top             =   8040
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   38
      Left            =   3810
      TabIndex        =   85
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   18
      afterdecimal    =   3
   End
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   39
      Left            =   5370
      TabIndex        =   87
      Top             =   8400
      Width           =   1335
      _ExtentX        =   2355
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   37
      Left            =   810
      TabIndex        =   83
      Top             =   8400
      Width           =   2895
      _ExtentX        =   5106
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
   Begin 合同管理.FTextBox txtSec 
      Height          =   300
      Index           =   40
      Left            =   6810
      TabIndex        =   89
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
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
      MaxLength       =   16
      afterdecimal    =   3
   End
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   2
      Left            =   450
      TabIndex        =   84
      Top             =   5520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   3
      Left            =   450
      TabIndex        =   86
      Top             =   5880
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   4
      Left            =   450
      TabIndex        =   88
      Top             =   6240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   5
      Left            =   450
      TabIndex        =   90
      Top             =   6600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   6
      Left            =   450
      TabIndex        =   91
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   7
      Left            =   450
      TabIndex        =   92
      Top             =   7320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   8
      Left            =   450
      TabIndex        =   93
      Top             =   7680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   9
      Left            =   450
      TabIndex        =   94
      Top             =   8040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdSubRow 
      Height          =   315
      Index           =   10
      Left            =   450
      TabIndex        =   95
      Top             =   8400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   1
      Left            =   8370
      TabIndex        =   100
      Top             =   5160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "^"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   2
      Left            =   8370
      TabIndex        =   101
      Top             =   5520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "^"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   3
      Left            =   8370
      TabIndex        =   102
      Top             =   5880
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   4
      Left            =   8370
      TabIndex        =   103
      Top             =   6240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   8
      Left            =   8370
      TabIndex        =   104
      Top             =   7680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   5
      Left            =   8370
      TabIndex        =   105
      Top             =   6600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   6
      Left            =   8370
      TabIndex        =   106
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   7
      Left            =   8370
      TabIndex        =   107
      Top             =   7320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   10
      Left            =   8370
      TabIndex        =   108
      Top             =   8400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdUp 
      Height          =   315
      Index           =   9
      Left            =   8370
      TabIndex        =   109
      Top             =   8040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   1
      Left            =   8730
      TabIndex        =   110
      Top             =   5160
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   2
      Left            =   8730
      TabIndex        =   111
      Top             =   5520
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   3
      Left            =   8730
      TabIndex        =   112
      Top             =   5880
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   4
      Left            =   8730
      TabIndex        =   113
      Top             =   6240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   8
      Left            =   8730
      TabIndex        =   114
      Top             =   7680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   5
      Left            =   8730
      TabIndex        =   115
      Top             =   6600
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   6
      Left            =   8730
      TabIndex        =   116
      Top             =   6960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   7
      Left            =   8730
      TabIndex        =   117
      Top             =   7320
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   10
      Left            =   8730
      TabIndex        =   118
      Top             =   8400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   9
      Left            =   8730
      TabIndex        =   119
      Top             =   8040
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin 合同管理.FTextBox txtSub 
      Height          =   300
      Index           =   15
      Left            =   1230
      TabIndex        =   9
      Top             =   4320
      Width           =   7755
      _ExtentX        =   13679
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
   Begin VB.Label Label6 
      Caption         =   "备    注"
      Height          =   300
      Left            =   450
      TabIndex        =   123
      Top             =   4380
      Width           =   720
   End
   Begin VB.Label Label5 
      Caption         =   "已 结 算"
      Height          =   255
      Left            =   5325
      TabIndex        =   121
      Top             =   3645
      Width           =   720
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
      Left            =   3570
      TabIndex        =   120
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lbl 
      Caption         =   "实际工作量(KM2)"
      Height          =   255
      Index           =   4
      Left            =   6855
      TabIndex        =   99
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label lbl 
      Caption         =   "合同单价(元)"
      Height          =   255
      Index           =   3
      Left            =   5490
      TabIndex        =   98
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Label lbl 
      Caption         =   "工作量(KM2)"
      Height          =   255
      Index           =   2
      Left            =   4050
      TabIndex        =   97
      Top             =   4920
      Width           =   1155
   End
   Begin VB.Label lbl 
      Caption         =   "工 作 内 容"
      Height          =   255
      Index           =   1
      Left            =   1770
      TabIndex        =   96
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "付款方式"
      Height          =   255
      Left            =   450
      TabIndex        =   45
      Top             =   4020
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "项目名称"
      Height          =   255
      Left            =   450
      TabIndex        =   44
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label Label33 
      Caption         =   "工程地点"
      Height          =   255
      Left            =   450
      TabIndex        =   43
      Top             =   3675
      Width           =   735
   End
   Begin VB.Label Label36 
      Caption         =   "承 揽 人"
      Height          =   255
      Left            =   450
      TabIndex        =   42
      Top             =   2205
      Width           =   735
   End
   Begin VB.Label Label38 
      Caption         =   "承包方式"
      Height          =   255
      Left            =   450
      TabIndex        =   41
      Top             =   1845
      Width           =   735
   End
   Begin VB.Label Label39 
      Caption         =   "预算借支金额(元)"
      Height          =   255
      Left            =   4650
      TabIndex        =   40
      Top             =   2850
      Width           =   1455
   End
   Begin VB.Label Label40 
      Caption         =   "结算价(元)"
      Height          =   255
      Left            =   5175
      TabIndex        =   39
      Top             =   1995
      Width           =   915
   End
   Begin VB.Label Label42 
      Caption         =   "进场人数"
      Height          =   255
      Left            =   450
      TabIndex        =   38
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label Label43 
      Caption         =   "录入日期"
      Height          =   255
      Left            =   5325
      TabIndex        =   37
      Top             =   3270
      Width           =   735
   End
   Begin VB.Label Label44 
      Caption         =   "合同编号"
      Height          =   255
      Left            =   450
      TabIndex        =   36
      Top             =   795
      Width           =   735
   End
   Begin VB.Label Label45 
      Caption         =   "项目编号"
      Height          =   255
      Left            =   450
      TabIndex        =   35
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label47 
      Caption         =   "合同名称"
      Height          =   255
      Left            =   5310
      TabIndex        =   34
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label50 
      Caption         =   "合同总价(元)"
      Height          =   255
      Left            =   4965
      TabIndex        =   33
      Top             =   1170
      Width           =   1080
   End
   Begin VB.Label Label51 
      Caption         =   "结算日期"
      Height          =   255
      Left            =   5325
      TabIndex        =   32
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label52 
      Caption         =   "退场日期"
      Height          =   255
      Left            =   450
      TabIndex        =   31
      Top             =   3300
      Width           =   735
   End
   Begin VB.Label Label53 
      Caption         =   "进场日期"
      Height          =   255
      Left            =   450
      TabIndex        =   30
      Top             =   2925
      Width           =   735
   End
   Begin VB.Label Label54 
      Caption         =   "其它[补贴...](元)"
      Height          =   255
      Left            =   4530
      TabIndex        =   29
      Top             =   1590
      Width           =   1545
   End
End
Attribute VB_Name = "frmInputSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dblBorrow As Double          ' 借支金额
Public rowHeight As Integer         '行高
Public rows As Byte                 '行数

Private Sub chkFinish_Click()
    cmdSave.Enabled = True
End Sub

Private Sub cmdAddRow_Click(Index As Integer)
    rows = rows + 1
    If rows > 10 Then
        rows = 10
        Exit Sub
    End If

    Me.Height = Me.Height + rowHeight
    cmdSave.Top = cmdSave.Top + rowHeight
    cmdExit.Top = cmdExit.Top + rowHeight
    
    For i = 1 To 4
        txtSec(4 * (rows - 1) + i).Visible = True
        txtSec(4 * (rows - 1) + i).Enabled = True
    Next
    cmdUp(rows).Visible = True
    cmdDown(rows).Visible = True
    cmdSubRow(rows).Visible = True
    
    cmdSave.Enabled = True
End Sub

Private Sub cmdCLS_Click(Index As Integer)
    txtSub(Index).Text = ""
End Sub

Private Sub cmdSubRow_Click(Index As Integer)
    Dim txt(4) As String
    If Index < rows Then
        For i = Index To rows - 1
            For j = 1 To 4
                txtSec((i - 1) * 4 + j).Text = txtSec(i * 4 + j).Text
            Next
        Next
        For i = 1 To 4
            txtSec((rows - 1) * 4 + i).Text = ""
        Next
    End If
        
    rows = rows - 1
    If rows < 1 Then
        rows = 1
        Exit Sub
    End If
    
    
    Me.Height = Me.Height - rowHeight
    cmdSave.Top = cmdSave.Top - rowHeight
    cmdExit.Top = cmdExit.Top - rowHeight
    
    
    
    
    For i = 1 To 4
        txtSec(rows * 4 + i).Visible = False
        txtSec(rows * 4 + i).Enabled = False
    Next
    
    cmdUp(rows + 1).Visible = False
    cmdDown(rows + 1).Visible = False
    cmdSubRow(rows + 1).Visible = False
    
    cmdSave.Enabled = True
    
End Sub

Private Sub cobCBFS_Expand()
    cmdSave.Enabled = True
End Sub
Private Sub DTPicker_CloseUp(Index As Integer)
    txtSub(Index).Text = DTPicker(Index).value
End Sub

Private Sub Form_Activate()
    
    If UCase(DataOperateState) = "EDIT" Then
         txtSub(1).SetFocus
    End If
    
End Sub

Private Sub Form_Load()
'On Error GoTo aaaa
    'Me.WindowState = vbMaximized    '最大化窗口
    Width = 9600
    Height = 9915
    Me.Top = 500
    Me.Left = 1500
    
    rowHeight = 360
    rows = 1
    
    IsBorrowEdit = False
    imgIcon.Picture = frmMain.cmdLeft(2).Picture
    Label28.caption = ""
    For i = 0 To UBound(strMode)
        cobCBFS.AddItem strMode(i, 0)
    Next
    cobCBFS.ListIndex = 0
    txtSub(16).Text = Now()   '录入日期
    
    '设置工作量小数位数
    For i = 0 To 9
        txtSec(i * 4 + 2).afterdecimal = bytAfterDec
        txtSec(i * 4 + 4).afterdecimal = bytAfterDec
    Next

    
    Select Case DataOperateState
        Case "EDIT"
            fillMain
            fillData
            cmdSave.Enabled = False
            
        Case "ADD"           '新增子合同记录
            Me.caption = "数据录入"
            Label1.caption = Me.caption
            fillMain
            
    End Select
    
    SettxtSecBox
    
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical

End Sub
Sub fillData()         '子合同信息
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim temp As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    strSQL = "select * from sub where id=" & subID
    rs.Open strSQL, Conn, 1, 1
    
    If rs.EOF Then Exit Sub    '无记录
    
    For i = 1 To rs.Fields.Count - 3
        If i = 3 Then     '处理第2个字段(承包方式)
            For k = 0 To UBound(strMode)
                If strMode(k, 1) = rs.Fields(i).value Then cobCBFS.ListIndex = k '0-再发包,1-自做
            Next
        Else
            If Not IsNull(rs.Fields(i).value) Then
                If Not FieldTypeIsChar(rs.Fields(i).Type) And rs.Fields(i).value = 0 Then    '数值型字段且值为零
                
                Else
                    txtSub(i).Text = rs.Fields(i).value
                End If
            End If
      
        End If
       
    Next
    
    chkFinish.value = IIf(rs("yjs"), 1, 0)
    
    rs.Close
    
    strSQL = "select * from subsec where zhtid=" & subID
    rs.Open strSQL, Conn, 1, 1
    If rs.EOF Then Exit Sub
    rows = 0
    Do While Not rs.EOF
        
        rows = rows + 1
        
        For i = 1 To 4
            
            If Not IsNull(rs.Fields(i).value) Then
                txtSec((rows - 1) * 4 + i).Text = rs.Fields(i).value
            End If
        
        Next
        rs.MoveNext
    
    Loop
    
    If rows = 0 Then rows = 1
    
    rs.Close
        
End Sub
Sub fillMain()
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    DBConnect
    strSQL = "select htbh,htmc from main where id=" & mainID
    rs.Open strSQL, Conn, 1, 1
    
    If rs.EOF Then Exit Sub  '无记录
    
    If Not IsNull(rs("htbh")) Then
        txtNo.Text = rs("htbh")      '合同编号
    Else
        txtNo.Text = "未录入"
    End If
    
    If Not IsNull(rs("htmc")) Then
        txtName.Text = rs("htmc")    '合同名称
    Else
        txtName.Text = "未录入"
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

Private Sub txtSec_Change(Index As Integer)
    cmdSave.Enabled = True
End Sub

Private Sub txtSec_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Index < 40 Then txtSec(Index + 1).SetFocus
    
End Sub

Private Sub txtSub_Change(Index As Integer)
    cmdSave.Enabled = True
    Label28.caption = ""
    Exit Sub
    Select Case Index
        Case 12, 17, 18
            If Trim(txtSub(12).Text) <> "" And Trim(txtSub(18).Text) <> "" Then
                temp = CDec(Trim(txtSub(12).Text)) * CDec(Trim(txtSub(18).Text))
                temp = Int(temp * 100 + 0.5) / 100   '四舍五入,保留二位小数
                txtSub(13).Text = Format(temp, "0.00")
                txtSub(19).Text = txtSub(13).Text
                If Trim(txtSub(15).Text) <> "" Then
                    txtSub(19).Text = CDec(txtSub(19).Text) + CDec(txtSub(15).Text)
                End If
            End If
        
    End Select
    
End Sub

Private Sub txtSub_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Index = 1 Or Index = 15 Then
            
        Else
            txtSub(Index + 1).SetFocus
        End If
        
    End If
End Sub

Private Sub cmdSave_Click()
  On Error GoTo errmsg
    
    If Trim(txtSub(1).Text) = "" Then
        MsgBox "未填写[项目编号]", vbInformation, Me.caption
        txtSub(1).SetFocus
        Exit Sub
    End If
    

        
    Select Case DataOperateState
        Case "EDIT"
            UpdateData "sub", 17, subID
            Form_Unload 0
        
        Case "ADD"
            AddRecord "sub", 17
            
            For i = 1 To 16
                If i <> 3 Then
                    txtSub(i).Text = ""
                End If
            Next
            For i = 1 To rows
                For j = 1 To 4
                    txtSec((i - 1) * 4 + j).Text = ""
                Next
            Next
            
            Width = 9600
            Height = 9555
            cmdSave.Top = 8565
            cmdExit.Top = cmdSave.Top
            rows = 1
            SettxtSecBox
            
            
            chkFinish.value = 0
            txtSub(16).Text = Now()
            cmdSave.Enabled = False
            Label28.caption = "保存成功!"
            txtSub(1).SetFocus
            
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
    strSQL = "xmbh='" & Trim(txtSub(1).Text) & "',cbfs=" & strMode(cobCBFS.ListIndex, 1)
    strSQL = strSQL & ",xmmc=" & IIf(Trim(txtSub(2).Text) = "", "NULL", "'" & Trim(txtSub(2).Text) & "'")
    For i = 4 To byteFieldCount - 1
        strText = Trim(txtSub(i).Text)
            If strText = "" Then
                strText = "Null"
            Else
                If FieldTypeIsChar(rs.Fields(i).Type) Then     '字符或日期字段,加单引号
                    strText = "'" & strText & "'"
                End If
            End If
                    
        strSQL = strSQL & "," & rs.Fields(i).Name & "=" & strText
    
    Next
    strSQL = strSQL & "," & rs.Fields(i).Name & "=" & IIf(chkFinish.value, "true", "false")
    strSQL = "update " & strTable & " set " & strSQL & " where id=" & lngID
    'MsgBox strSQL
    Conn.Execute strSQL
    
    
    '更新SubSec表
    strSQL = "delete from subsec where zhtid=" & lngID
    Conn.Execute strSQL
    
    
    SaveSubSec subID
        
    Unload Me
    
    Exit Sub
End Sub
Sub AddRecord(strTable As String, byteFieldCount As Byte)    'flag -借支记录标志
    Dim strField, strValue, strText, strSQL As String
    Dim i, iCount As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    DBConnect
    strSQL = "select top 1  * from " & strTable
    
    rs.Open strSQL, Conn, 1, 1
    
    strField = "(xmbh,xmmc,cbfs"
    strValue = "('" & Trim(txtSub(1).Text) & "','" & Trim(txtSub(2).Text) & "'," & strMode(cobCBFS.ListIndex, 1)
    For i = 4 To byteFieldCount - 1
        strText = Trim(txtSub(i).Text)
                        
        If strText <> "" Then
            
            strField = strField & "," & rs.Fields(i).Name
            
            If FieldTypeIsChar(rs.Fields(i).Type) Then   '字符或日期字段,加单引号
                strText = "'" & strText & "'"
            End If
         
            strValue = strValue & "," & strText
            
        End If
        
    Next
    strField = strField & "," & rs.Fields(i).Name
    strValue = strValue & "," & IIf(chkFinish.value, "true", "false")
    strField = strField & ",zhtid" & ")"
    strValue = strValue & "," & mainID & ")"
    
    strSQL = "insert into " & strTable & " " & strField & "  values" & strValue
    Conn.Execute strSQL
    
    rs.Close
    
    strSQL = "select top 1 id from sub order by id desc"
    rs.Open strSQL, Conn, 1, 1
    
    If rs.EOF Then Exit Sub
    
    SaveSubSec rs("id")    '增加新记录到子合同二SubSec表

End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Sub SaveSubSec(id As Long)
    Dim strField, strValue, strSQL As String
    strField = "("
    strValue = "("
    
    For i = 1 To rows
        For j = 1 To 4
            If Trim(txtSec((i - 1) * 4 + j).Text) <> "" Then
                Exit For
            End If
        Next
        If j < 5 Then
            
            strSQL = "insert into subsec(gzny,gzl,htdj,sjgzl,lrrq,zhtid) values("
            strSQL = strSQL & "'" & Trim(txtSec((i - 1) * 4 + 1).Text) & "'," & _
                                  "0" & Trim(txtSec((i - 1) * 4 + 2).Text) & "," & _
                                  "0" & Trim(txtSec((i - 1) * 4 + 3).Text) & "," & _
                                  "0" & Trim(txtSec((i - 1) * 4 + 4).Text) & "," & _
                                  "'" & Now() & "'," & id & ")"
            
        'MsgBox strSQL
        Conn.Execute strSQL
        End If
    
    Next

End Sub

Sub SettxtSecBox()
    
    For i = rows + 1 To 10
        For j = 1 To 4
            txtSec((i - 1) * 4 + j).Visible = False
            txtSec((i - 1) * 4 + j).Enabled = False
        Next
        cmdUp(i).Visible = False
        cmdDown(i).Visible = False
        cmdSubRow(i).Visible = False
    Next
    
    cmdSave.Top = cmdSave.Top - rowHeight * (10 - rows)
    cmdExit.Top = cmdSave.Top
    Me.Height = Me.Height - rowHeight * (10 - rows)

End Sub
