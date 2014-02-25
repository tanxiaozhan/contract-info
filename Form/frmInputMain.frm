VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputMain 
   BackColor       =   &H00D4B89D&
   Caption         =   "数据修改"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   9585
   Begin VB.CheckBox chkFinish 
      BackColor       =   &H00D4B89D&
      Height          =   255
      Left            =   6480
      TabIndex        =   18
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "合 同 信 息"
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
         Left            =   4080
         TabIndex        =   122
         Top             =   120
         Width           =   1260
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
      Left            =   5580
      TabIndex        =   91
      Top             =   9165
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
      Left            =   2940
      TabIndex        =   89
      Top             =   9165
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
   Begin 合同管理.XPButton cmdCLS 
      Height          =   330
      Index           =   10
      Left            =   8265
      TabIndex        =   21
      Top             =   690
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
      Index           =   14
      Left            =   8025
      TabIndex        =   22
      Top             =   2400
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   112525313
      CurrentDate     =   39889
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   300
      Index           =   11
      Left            =   8025
      TabIndex        =   23
      Top             =   1155
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   112525313
      CurrentDate     =   39889
   End
   Begin MSComCtl2.DTPicker DTPicker 
      Height          =   300
      Index           =   10
      Left            =   8025
      TabIndex        =   24
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Format          =   112525313
      CurrentDate     =   39889
   End
   Begin 合同管理.XPButton cmdCLS 
      Height          =   330
      Index           =   11
      Left            =   8265
      TabIndex        =   25
      Top             =   1125
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
      Index           =   14
      Left            =   8265
      TabIndex        =   26
      Top             =   2370
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   2
      Left            =   1935
      TabIndex        =   1
      Top             =   1080
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   3
      Left            =   1935
      TabIndex        =   2
      Top             =   1440
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   4
      Left            =   1935
      TabIndex        =   3
      Top             =   1800
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   5
      Left            =   1935
      TabIndex        =   4
      Top             =   2160
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   6
      Left            =   1935
      TabIndex        =   5
      Top             =   2520
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   10
      Left            =   6435
      TabIndex        =   11
      Top             =   720
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   11
      Left            =   6435
      TabIndex        =   12
      Top             =   1155
      Width           =   2445
      _ExtentX        =   4313
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
      Locked          =   -1  'True
   End
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   12
      Left            =   6435
      TabIndex        =   13
      Top             =   1575
      Width           =   2445
      _ExtentX        =   4313
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
   End
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   13
      Left            =   6435
      TabIndex        =   14
      Top             =   1995
      Width           =   2445
      _ExtentX        =   4313
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
   End
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   14
      Left            =   6435
      TabIndex        =   15
      Top             =   2400
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FCombo cobType 
      Height          =   300
      Left            =   1935
      TabIndex        =   0
      Top             =   720
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   7
      Left            =   1935
      TabIndex        =   6
      Top             =   2880
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   8
      Left            =   1935
      TabIndex        =   7
      Top             =   3240
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   9
      Left            =   1935
      TabIndex        =   8
      Top             =   3600
      Width           =   2445
      _ExtentX        =   4313
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
   End
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   15
      Left            =   6435
      TabIndex        =   16
      Top             =   2820
      Width           =   2445
      _ExtentX        =   4313
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
   End
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   18
      Left            =   6435
      TabIndex        =   17
      Top             =   3225
      Width           =   2445
      _ExtentX        =   4313
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   16
      Left            =   1935
      TabIndex        =   9
      Top             =   3960
      Width           =   6930
      _ExtentX        =   12224
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
      Left            =   465
      TabIndex        =   45
      Top             =   5400
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
      Left            =   3780
      TabIndex        =   47
      Top             =   5430
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
      Left            =   5340
      TabIndex        =   48
      Top             =   5430
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
      Left            =   780
      TabIndex        =   46
      Top             =   5430
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
      Left            =   6765
      TabIndex        =   49
      Top             =   5430
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
      Left            =   3780
      TabIndex        =   51
      Top             =   5790
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
      Left            =   5340
      TabIndex        =   52
      Top             =   5790
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
      Left            =   780
      TabIndex        =   50
      Top             =   5790
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
      Left            =   6765
      TabIndex        =   53
      Top             =   5790
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
      Left            =   3780
      TabIndex        =   55
      Top             =   6150
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
      Left            =   5340
      TabIndex        =   56
      Top             =   6150
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
      Left            =   780
      TabIndex        =   54
      Top             =   6150
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
      Left            =   6780
      TabIndex        =   57
      Top             =   6150
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
      Left            =   3780
      TabIndex        =   59
      Top             =   6510
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
      Left            =   5340
      TabIndex        =   60
      Top             =   6510
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
      Left            =   780
      TabIndex        =   58
      Top             =   6510
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
      Left            =   6780
      TabIndex        =   61
      Top             =   6510
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
      Left            =   3780
      TabIndex        =   63
      Top             =   6870
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
      Left            =   5340
      TabIndex        =   64
      Top             =   6870
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
      Left            =   780
      TabIndex        =   62
      Top             =   6870
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
      Left            =   6780
      TabIndex        =   65
      Top             =   6870
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
      Left            =   3780
      TabIndex        =   67
      Top             =   7230
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
      Left            =   5340
      TabIndex        =   68
      Top             =   7230
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
      Left            =   780
      TabIndex        =   66
      Top             =   7230
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
      Left            =   6780
      TabIndex        =   69
      Top             =   7230
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
      Left            =   3780
      TabIndex        =   71
      Top             =   7590
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
      Left            =   5340
      TabIndex        =   72
      Top             =   7590
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
      Left            =   795
      TabIndex        =   70
      Top             =   7590
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
      Left            =   6780
      TabIndex        =   73
      Top             =   7590
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
      Left            =   3780
      TabIndex        =   75
      Top             =   7950
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
      Left            =   5340
      TabIndex        =   76
      Top             =   7950
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
      Left            =   780
      TabIndex        =   74
      Top             =   7950
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
      Left            =   6780
      TabIndex        =   77
      Top             =   7950
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
      Left            =   3780
      TabIndex        =   79
      Top             =   8310
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
      Left            =   5340
      TabIndex        =   80
      Top             =   8310
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
      Left            =   780
      TabIndex        =   78
      Top             =   8310
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
      Left            =   6780
      TabIndex        =   81
      Top             =   8310
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
      Left            =   3780
      TabIndex        =   83
      Top             =   8670
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
      Left            =   5340
      TabIndex        =   85
      Top             =   8670
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
      Left            =   780
      TabIndex        =   82
      Top             =   8670
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
      Left            =   6780
      TabIndex        =   87
      Top             =   8670
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
      Left            =   465
      TabIndex        =   84
      Top             =   5760
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
      Left            =   465
      TabIndex        =   86
      Top             =   6120
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
      Left            =   465
      TabIndex        =   88
      Top             =   6480
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
      Left            =   465
      TabIndex        =   90
      Top             =   6840
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
      Left            =   465
      TabIndex        =   92
      Top             =   7200
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
      Left            =   465
      TabIndex        =   93
      Top             =   7560
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
      Left            =   465
      TabIndex        =   94
      Top             =   7920
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
      Left            =   465
      TabIndex        =   95
      Top             =   8280
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
      Left            =   465
      TabIndex        =   96
      Top             =   8640
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
      Left            =   8295
      TabIndex        =   97
      Top             =   5400
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
      Left            =   8295
      TabIndex        =   98
      Top             =   5760
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
      Left            =   8295
      TabIndex        =   99
      Top             =   6120
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
      Index           =   4
      Left            =   8295
      TabIndex        =   100
      Top             =   6480
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
      Index           =   8
      Left            =   8295
      TabIndex        =   101
      Top             =   7920
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
      Index           =   5
      Left            =   8295
      TabIndex        =   102
      Top             =   6840
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
      Index           =   6
      Left            =   8295
      TabIndex        =   103
      Top             =   7200
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
      Index           =   7
      Left            =   8295
      TabIndex        =   104
      Top             =   7560
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
      Index           =   10
      Left            =   8295
      TabIndex        =   105
      Top             =   8640
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
      Index           =   9
      Left            =   8295
      TabIndex        =   106
      Top             =   8280
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
   Begin 合同管理.XPButton cmdDown 
      Height          =   315
      Index           =   1
      Left            =   8610
      TabIndex        =   107
      Top             =   5400
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
      Left            =   8610
      TabIndex        =   108
      Top             =   5760
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
      Left            =   8610
      TabIndex        =   109
      Top             =   6120
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
      Left            =   8610
      TabIndex        =   110
      Top             =   6480
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
      Left            =   8610
      TabIndex        =   111
      Top             =   7920
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
      Left            =   8610
      TabIndex        =   112
      Top             =   6840
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
      Left            =   8610
      TabIndex        =   113
      Top             =   7200
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
      Left            =   8610
      TabIndex        =   114
      Top             =   7560
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
      Left            =   8610
      TabIndex        =   115
      Top             =   8640
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
      Left            =   8610
      TabIndex        =   116
      Top             =   8280
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   17
      Left            =   1935
      TabIndex        =   10
      Top             =   4305
      Width           =   6930
      _ExtentX        =   12224
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
   Begin 合同管理.FTextBox txtBox 
      Height          =   300
      Index           =   19
      Left            =   1935
      TabIndex        =   124
      Top             =   4665
      Width           =   6915
      _ExtentX        =   12197
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
   Begin VB.Label Label18 
      BackColor       =   &H00D4B89D&
      Caption         =   "项目负责人"
      Height          =   225
      Left            =   1005
      TabIndex        =   125
      Top             =   4710
      Width           =   930
   End
   Begin VB.Label Label11 
      BackColor       =   &H00D4B89D&
      Caption         =   "备   注"
      Height          =   255
      Left            =   1275
      TabIndex        =   123
      Top             =   4330
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00D4B89D&
      Caption         =   "已 结 算"
      Height          =   255
      Left            =   5685
      TabIndex        =   121
      Top             =   3645
      Width           =   720
   End
   Begin VB.Label lbl 
      BackColor       =   &H00D4B89D&
      Caption         =   "分 项 内 容"
      Height          =   180
      Index           =   1
      Left            =   1620
      TabIndex        =   120
      Top             =   5150
      Width           =   1080
   End
   Begin VB.Label lbl 
      BackColor       =   &H00D4B89D&
      Caption         =   "工作量(KM2)"
      Height          =   180
      Index           =   2
      Left            =   4020
      TabIndex        =   119
      Top             =   5150
      Width           =   1155
   End
   Begin VB.Label lbl 
      BackColor       =   &H00D4B89D&
      Caption         =   "合同单价(元)"
      Height          =   180
      Index           =   3
      Left            =   5460
      TabIndex        =   118
      Top             =   5150
      Width           =   1095
   End
   Begin VB.Label lbl 
      BackColor       =   &H00D4B89D&
      Caption         =   "实际工作量(KM2)"
      Height          =   180
      Index           =   4
      Left            =   6780
      TabIndex        =   117
      Top             =   5150
      Width           =   1380
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D4B89D&
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
      Left            =   4275
      TabIndex        =   44
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D4B89D&
      Caption         =   "付款方式"
      Height          =   255
      Left            =   1185
      TabIndex        =   43
      Top             =   4005
      Width           =   735
   End
   Begin VB.Label Label17 
      BackColor       =   &H00D4B89D&
      Caption         =   "其他[补贴...](元)"
      Height          =   255
      Left            =   4905
      TabIndex        =   42
      Top             =   1620
      Width           =   1545
   End
   Begin VB.Label Label16 
      BackColor       =   &H00D4B89D&
      Caption         =   "进场日期"
      Height          =   255
      Left            =   5670
      TabIndex        =   41
      Top             =   765
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00D4B89D&
      Caption         =   "退场日期"
      Height          =   255
      Left            =   5670
      TabIndex        =   40
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label14 
      BackColor       =   &H00D4B89D&
      Caption         =   "结算日期"
      Height          =   255
      Left            =   5670
      TabIndex        =   39
      Top             =   2475
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00D4B89D&
      Caption         =   "合同总价(元)"
      Height          =   255
      Left            =   825
      TabIndex        =   38
      Top             =   3610
      Width           =   1080
   End
   Begin VB.Label Label8 
      BackColor       =   &H00D4B89D&
      Caption         =   "委托单位联系电话"
      Height          =   255
      Left            =   465
      TabIndex        =   37
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00D4B89D&
      Caption         =   "合同名称"
      Height          =   255
      Left            =   1185
      TabIndex        =   36
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00D4B89D&
      Caption         =   "委托单位联系人"
      Height          =   255
      Left            =   645
      TabIndex        =   35
      Top             =   1845
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackColor       =   &H00D4B89D&
      Caption         =   "合同编号"
      Height          =   255
      Left            =   1185
      TabIndex        =   34
      Top             =   1150
      Width           =   735
   End
   Begin VB.Label Label27 
      BackColor       =   &H00D4B89D&
      Caption         =   "录入日期"
      Height          =   255
      Left            =   5670
      TabIndex        =   33
      Top             =   3300
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00D4B89D&
      Caption         =   "结余金额(元)"
      Height          =   255
      Left            =   5325
      TabIndex        =   32
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label19 
      BackColor       =   &H00D4B89D&
      Caption         =   "结算价(元)"
      Height          =   255
      Left            =   5505
      TabIndex        =   31
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label Label12 
      BackColor       =   &H00D4B89D&
      Caption         =   "合同类型"
      Height          =   255
      Left            =   1185
      TabIndex        =   30
      Top             =   780
      Width           =   735
   End
   Begin VB.Label Label9 
      BackColor       =   &H00D4B89D&
      Caption         =   "测绘内容"
      Height          =   255
      Left            =   1185
      TabIndex        =   29
      Top             =   3285
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H00D4B89D&
      Caption         =   "工程地点"
      Height          =   255
      Left            =   1185
      TabIndex        =   28
      Top             =   2925
      Width           =   735
   End
   Begin VB.Label Label37 
      BackColor       =   &H00D4B89D&
      Caption         =   "委托单位"
      Height          =   255
      Left            =   1185
      TabIndex        =   27
      Top             =   1485
      Width           =   735
   End
End
Attribute VB_Name = "frmInputMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    txtBox(Index).Text = ""
End Sub

Private Sub cobCBFS_Expand()
    cmdSave.Enabled = True
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

Private Sub cobType_Expand()
    cmdSave.Enabled = True

End Sub

Private Sub DTPicker_CloseUp(Index As Integer)
    txtBox(Index).Text = DTPicker(Index).value
End Sub

Private Sub Form_Activate()
     txtBox(2).SetFocus
End Sub

Private Sub Form_Load()
'On Error GoTo aaaa
    'Me.WindowState = vbMaximized    '最大化窗口
    Width = 9705
    Height = 10250
    Me.Top = 500
    Me.Left = 1500
    IsBorrowEdit = False
    imgIcon.Picture = frmMain.cmdLeft(2).Picture
    Label28.caption = ""
    For i = 0 To UBound(strContractType)
        cobType.AddItem strContractType(i, 0)
    Next
    cobType.ListIndex = 0
    txtBox(18).Text = Now()   '录入日期
    
    '设置工作量小数位数
    For i = 0 To 9
        txtSec(i * 4 + 2).afterdecimal = bytAfterDec
        txtSec(i * 4 + 4).afterdecimal = bytAfterDec
    Next
    
    rows = 1
    rowHeight = 360
    
    
    Select Case DataOperateState
        Case "EDIT"      '合同记录
            fillData
            cmdSave.Enabled = False
        
        Case "ADD"           '新增合同记录
            Me.caption = "数据录入"
            Label1.caption = Me.caption
    End Select
    
    SettxtSecBox
    
    
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical

End Sub
Sub fillData()         '合同信息
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim temp As String
    
    Set rs = New ADODB.Recordset
    DBConnect
    strSQL = "select * from main where id=" & mainID
    rs.Open strSQL, Conn, 1, 1
    If Not rs.EOF Then
        
        For i = 0 To UBound(strContractType)
            If strContractType(i, 1) = rs("htlx") Then cobType.ListIndex = i '0-公司,1-外协
        Next
    
        For i = 2 To rs.Fields.Count - 3
                If Trim(rs.Fields(i).value) <> "" Then
                    If Not FieldTypeIsChar(rs.Fields(i).Type) And rs.Fields(i).value = 0 Then
                    
                    Else
                        txtBox(i).Text = rs.Fields(i).value
                    End If
                End If
        
        Next
        
        chkFinish.value = IIf(rs("yjs"), 1, 0)
        If Trim(rs.Fields(rs.Fields.Count - 1).value) <> "" Then txtBox(19).Text = rs.Fields(rs.Fields.Count - 1).value
        
    End If
    
    rs.Close
    
    strSQL = "select * from mainsec where zhtid=" & mainID
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
Private Sub Form_Resize()
On Error Resume Next
    
    'PicTop.Width = Width * 0.5
    'PicTop.ScaleWidth = Me.ScaleWidth
    Cls
    'Line (2, 2)-(Width - 200, Height - 100), 10921638, B

End Sub

Private Sub Form_Unload(Cancel As Integer)
    DataOperateState = "ADD"
    frmMain.cmdLeft_Click 1  '显示列表
End Sub
Private Sub txtBox_Change(Index As Integer)
    cmdSave.Enabled = True
    Label28.caption = ""
    Exit Sub
    Select Case Index
        Case 12, 17, 18
            If Trim(txtBox(12).Text) <> "" And Trim(txtBox(18).Text) <> "" Then
                temp = CDec(Trim(txtBox(12).Text)) * CDec(Trim(txtBox(18).Text))
                temp = Int(temp * 100 + 0.5) / 100   '四舍五入,保留二位小数
                txtBox(13).Text = Format(temp, "0.00")
                txtBox(17).Text = txtBox(13).Text
                If Trim(txtBox(17).Text) <> "" Then
                    txtBox(17).Text = CDec(txtBox(17).Text) + CDec(txtBox(17).Text)
                End If
            End If
        
    End Select
    
End Sub

Private Sub txtBox_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Index < 17 Then
            txtBox(Index + 1).SetFocus
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errmsg
    
    If Trim(txtBox(2).Text) = "" Then
        MsgBox "未填写[合同编号]", vbInformation, Me.caption
        txtBox(2).SetFocus
        Exit Sub
    End If
    

        
    Select Case DataOperateState
        Case "EDIT"
            UpdateData "main", 19, mainID
            Form_Unload 0
        
        Case "ADD"
            AddRecord "main", 19
            
            For i = 2 To 18
                txtBox(i).Text = ""
            Next
            
            For i = 1 To rows
                For j = 1 To 4
                    txtSec((i - 1) * 4 + j).Text = ""
                Next
            Next
            Width = 9705
            Height = 9600
            cmdSave.Top = 8655
            cmdExit.Top = cmdSave.Top
            
            rows = 1
            
            SettxtSecBox
            
            chkFinish.value = 0
            txtBox(18).Text = Now()
            chkFinish.value = 0
            cmdSave.Enabled = False
            Label28.caption = "保存成功!"
            txtBox(2).SetFocus
            
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
    
    strSQL = "htlx=" & strContractType(cobType.ListIndex, 1)     '合同类型字段
    strSQL = strSQL & ",yjs=" & IIf(chkFinish.value, "true", "false")
    
    For i = 2 To byteFieldCount - 1
        strText = Trim(txtBox(i).Text)
        If strText = "" Then
            strText = "Null"
        Else
           If FieldTypeIsChar(rs.Fields(i).Type) Then     '字符或日期字段,加单引号
               strText = "'" & strText & "'"
           End If
        End If
           
        strSQL = strSQL & "," & rs.Fields(i).Name & "=" & strText
                    
    Next
    
    strText = "'" & Trim(txtBox(19).Text) & "'"     '项目负责人
    If strText = "''" Then strText = "Null"
    
    strSQL = strSQL & ",fzr=" & strText
    
    strSQL = "update " & strTable & " set " & strSQL & " where id=" & lngID
    'MsgBox strSQL
    Conn.Execute strSQL
            
            
    '更新mainSec表
    strSQL = "delete from mainsec where zhtid=" & lngID
    Conn.Execute strSQL
    
    
    SavemainSec lngID
        
            
    Form_Unload 0
    Exit Sub



End Sub
Sub AddRecord(strTable As String, byteFieldCount As Byte)    'flag -借支记录标志
    Dim strField, strValue, strText, strSQL As String
    Dim i, iCount As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    strSQL = "select top 1  * from " & strTable
    rs.Open strSQL, Conn, 1, 1
    
    strField = "(" & rs.Fields(1).Name
    strValue = "(" & strContractType(cobType.ListIndex, 1)
    
    For i = 2 To byteFieldCount - 1
        strText = Trim(txtBox(i).Text)
                        
        If strText <> "" Then
            strField = strField & "," & rs.Fields(i).Name
            
            If FieldTypeIsChar(rs.Fields(i).Type) Then   '字符或日期字段,加单引号
                strText = "'" & strText & "'"
            End If
         
            strValue = strValue & "," & strText
        End If
    Next
    strField = strField & "," & rs.Fields(i).Name & ",fzr)"
    strValue = strValue & "," & IIf(chkFinish.value, "true", "false") & ",'" & Trim(txtBox(19).Text) & "')"
    
    strSQL = "insert into " & strTable & " " & strField & "  values" & strValue
    
    Conn.Execute strSQL

    rs.Close
    
    strSQL = "select top 1 id from main order by id desc"
    rs.Open strSQL, Conn, 1, 1
    
    If rs.EOF Then Exit Sub
    
    SavemainSec rs("id")    '增加新记录到合同二mainSec表


End Sub
Private Sub cmdExit_Click()
    Form_Unload 0
End Sub

Private Sub txtSec_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    cmdSave.Enabled = True

End Sub
Sub SavemainSec(id As Long)
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
            
            strSQL = "insert into mainsec(fxny,gzl,htdj,sjgzl,lrrq,zhtid) values("
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

