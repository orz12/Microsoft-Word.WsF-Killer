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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton CmdAbout 
      Caption         =   "����"
      Height          =   855
      Left            =   6600
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "���ǰ��ֹcmd��cscript��wscript����������"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�ָ��ļ�"
      Enabled         =   0   'False
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
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
MsgBox "����������UI�����þ��У���ϵ��ʽtemp_lyq@163.com����ӭ������Spring In Pink��һ�����ӡ�", vbInformation
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Check1.Value = 1 Then
Me.Caption = "������ֹ���̲�������ʱ��������"
AutoTerminateProcess
End If
Me.Caption = "������ͼɾ�����������Ȩ�޹�һ�㶼û�����⡭��"
Dim WshShell, bKey
Set WshShell = WScript.CreateObject("WScript.Shell")
'ɾ��ע���
WshShell.RegDelete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Microsoft Word"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\Microsoft Word"
Set WshShell = Nothing
Me.Caption = "ɾ��������Ȩ�޹�һ�㶼û�����⣬��Ҳ��˵����"
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
If MsgBox("�˲�������һ��Ӱ�죺" & vbCrLf & "1������Ϊϵͳ�����ص��ļ����ܻ���ʾ����������ƽʱ����һЩ����������������" & vbCrLf & "2�������ļ��м������ļ��������еĿ�ݷ�ʽ���ᱻ�޲��ֱ��ɾ��������̫���ˣ�" & vbCrLf & "�Ƿ�ִ�У�", vbInformation Or vbYesNo, "���ù������㡾�ǡ��Ϳ��ԣ���һ����кڿ��@_@����ʧ�˾ͺ���") = vbNo Then Exit Sub
Shell "cmd.exe /c attrib *.* -s -h /s /d"
Shell "cmd.exe /c del *.lnk /f /s /q"

End Sub

Private Sub Form_Load()
On Error Resume Next

    If App.LogMode = 0 Then Stop
    If IsDebuggerPresent Then MsgBox "���򱻵���,����ȷ���˳���", vbSystemModal Or vbCritical: End

    If MsgBox("����Ŀ�����ȷ����Ϊ��ͬhu��lve�������" & vbCrLf & "1������Ʒ��ֹ���ã������ֱ�Ӻ�/������ʧ�����߸Ų�����:P��ע�ⱸ������" & vbCrLf & "2�����ڸ�Ŀ¼����!!" & vbCrLf & "3����Ҫ����ԱȨ��XD", vbSystemModal Or vbExclamation Or vbOKCancel) = vbCancel Then Unload Me
    If RtlAdjustPrivilege(20, 1, 0, 0) = &HC000007C Then RtlAdjustPrivilege 20, 0, 0, 0 '�򿪽���Ȩ��
    
End Sub

Private Sub Form_Terminate()
    Shell "cmd.exe /c attrib " & App.Path & "Microsoft Word.WsF" & " +s +h"
End Sub
