VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00FDFDFD&
   Caption         =   "欢迎使用本系统"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   ControlBox      =   0   'False
   Icon            =   "frmWelcome.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   438
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   598
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'超市销售系统
'程序开发：lc_mtt
'CSDN博客：http://blog.csdn.net/lc_mtt/
'个人主页：http://www.3lsoft.com
'邮箱：3lsoft@163.com
'注：此代码禁止用于商业用途。有修改者发我一份，谢谢！
'---------------- 开源世界，你我更进步 ----------------

Private Sub Form_Load()
    Me.WindowState = 2
End Sub

Private Sub Form_Resize()
On Error Resume Next
    imgBack.Move (Width / 15 - imgBack.Width) / 2, (Height / 15 - imgBack.Height) / 2
End Sub
