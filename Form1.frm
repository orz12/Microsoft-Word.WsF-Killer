VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Microsoft Word.WsF Killer v2.718"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   8100
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CmdAbout 
      Caption         =   "关于"
      Height          =   855
      Left            =   6600
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "清除前终止cmd、cscript与wscript并予以锁定"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "恢复文件"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尝试清除"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   975
      Left            =   1080
      TabIndex        =   4
      Top             =   2520
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal NewValue&, ByVal NewThread&, Oldvalue&)


Private Sub CmdAbout_Click()
MsgBox "作者懒得做UI，能用就行；联系方式temp_lyq@163.com，欢迎交流；Spring In Pink是一首曲子。", vbInformation
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Check1.Value = 1 Then
Me.Caption = "正在终止进程并尝试暂时锁定……"
AutoTerminateProcess
End If
Me.Caption = "正在试图删除开机启动项，权限够一般都没有问题……"
Dim WshShell, bKey
Set WshShell = WScript.CreateObject("WScript.Shell")
'删除注册表
WshShell.RegDelete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Microsoft Word"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Microsoft Word"
Set WshShell = Nothing
Me.Caption = "删本体辣，权限够一般都没有问题，有也不说……"
Kill App.Path & "Microsoft Word.WsF"
MkDir App.Path & "Microsoft Word.WsF"
bKey = GetAttr(App.Path & "Microsoft Word.WsF")
SetAttr f, bKey Or (vbReadOnly + vbHidden + vbSystem)
Dim oShell
Dim strHomeFolder As String
'Get folder
Set oShell = CreateObject("WScript.Shell")
strHomeFolder = oShell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft Office\"
strHomeFolder = oShell.ExpandEnvironmentStrings(strHomeFolder)
Set oShell = Nothing

Command2.Enabled = True
Me.Caption = "Microsoft Word.WsF Killer v2.718"
MsgBox "bingo!", vbInformation, "CoooolKiller"
End Sub

Private Sub Command2_Click()
On Error Resume Next
If MsgBox("此操作会有一定影响：" & vbCrLf & "1、属性为系统、隐藏的文件可能会显示，看起来比平时多了一些东西属于正常现象" & vbCrLf & "2、所在文件夹及其子文件夹下所有的快捷方式均会被无差别直接删除（作者太懒了）" & vbCrLf & "是否执行？", vbInformation Or vbYesNo, "不用管它，点【是】就可以；过一会会有黑框框@_@，消失了就好了") = vbNo Then Exit Sub
Shell "cmd.exe /c attrib *.* -s -h /s /d"
Shell "cmd.exe /c del *.lnk /f /s /q"

End Sub

Private Sub Form_Load()
On Error Resume Next

    If App.LogMode = 0 Then Stop
    If IsDebuggerPresent Then MsgBox "程序被调试,单击确定退出。", vbSystemModal Or vbCritical: End

    If MsgBox("【醒目】点击确定视为认同hu意lve以下条款：" & vbCrLf & "1、本产品禁止商用，所造成直接和/或间接损失开发者概不负责:P，注意备份数据" & vbCrLf & "2、请于根目录运行!!" & vbCrLf & "3、需要管理员权限XD", vbSystemModal Or vbExclamation Or vbOKCancel) = vbCancel Then Unload Me
    If RtlAdjustPrivilege(20, 1, 0, 0) = &HC000007C Then RtlAdjustPrivilege 20, 0, 0, 0 '打开进程权限
    
End Sub

Private Sub Form_Terminate()
    Shell "cmd.exe /c attrib " & App.Path & "Microsoft Word.WsF" & " +s +h"
End Sub
