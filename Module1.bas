Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPheaplist = &H1
Private Const TH32CS_SNAPthread = &H4
Private Const TH32CS_SNAPmodule = &H8
Private Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Private Const MAX_PATH As Integer = 260
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
Private Const GENERIC_EXECUTE = &H20000000
Private Const OPEN_ALWAYS = 4
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Private Const PROCESS_DUP_HANDLE As Long = &H40
Private Const INVALID_HANDLE_VALUE = -1
Private Const DUPLICATE_SAME_ACCESS = &H2
Private Const DUPLICATE_CLOSE_SOURCE = &H1

Dim xxNull As SECURITY_ATTRIBUTES
Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Sub AutoTerminateProcess()
    Dim proc As PROCESSENTRY32
    Dim snap As Long
    Dim exeName As String
    Dim theloop As Long
    Dim hd As Long
    Dim HwndToExePath As String
    snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) '获得进程“快照”的句柄
    proc.dwSize = Len(proc)
    theloop = ProcessFirst(snap, proc) '获取第一个进程，并得到其返回值
    While theloop <> 0 '当返回值非零时继续获取下一个进程
        exeName = LCase(proc.szExeFile)
        If (Left(exeName, 7) = "cmd.exe" Or Left(exeName, 11) = "cscript.exe" Or Left(exeName, 11) = "wscript.exe") Then
        hd = OpenProcess(PROCESS_ALL_ACCESS, True, proc.th32ProcessID)
            If hd <> 0 Then
                'l = GetModuleFileNameEx(hd, 0, HwndToExePath, 255)
                Dim cbNeeded As Long
                Dim szBuf(1 To 250) As Long
                Dim Ret As Long
                Dim szPathName As String
                Dim nSize As Long
                Ret = EnumProcessModules(hd, szBuf(1), 250, cbNeeded)
                'If Ret <> 0 Then
                    szPathName = Space(260)
                    nSize = 500
                    Ret = GetModuleFileNameEx(hd, szBuf(1), szPathName, nSize)
                    HwndToExePath = Left(szPathName, Ret)
                'End If
                HwndToExePath = Trim(HwndToExePath)
                TerminateProcess hd, 0
                If HwndToExePath <> "" Then CreateFile HwndToExePath, 0, 0, xxNull, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0 '打开文件(获得文件句柄)
                CloseHandle hd
            End If
        End If
        theloop = ProcessNext(snap, proc)
    Wend
    CloseHandle snap '关闭进程“快照”句柄
End Sub



