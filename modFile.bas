Attribute VB_Name = "modFile"
''==============================================================
''
''      modFiles [modfiles.bas]
''
''==============================================================


Option Explicit

''----------------''

'--------------------------------------------------------
' User variable
'--------------------------------------------------------
Global LogName As String     'scan-wip
Public TCnt    As Integer
Public Ccnt    As Integer
'--------------------------------------------------------


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


''----------------''
Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
      (ByVal hwnd As Integer, ByVal wMsg As Integer, _
      ByVal wParam As Integer, lParam As Any) As Integer

Public Const WM_CLOSE = &H10
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101


Public m_sEXEName As String
Public m_sDosCaption As String

''----------------''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''[ for WaitForProcess() ]
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
''
Public Const INFINITE = &HFFFF     ' dwMilliseconds parameter
'''''''''''''''''''''''''''''''''''' WaitForSingleObject return values:
Public Const WAIT_TIMEOUT = &H102
Public Const WAIT_FAILED = &HFFFFFFFF
Public Const WAIT_OBJECT_0 = &H0
Public Const STILL_ACTIVE = &H103

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, _
                        ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, _
                        lpExitCode As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, _
                        ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''']

''----------------''
Public Declare Function GetFileAttributes Lib "kernel32" Alias _
                        "GetFileAttributesA" (ByVal lpFileName As String) As Long
                        

''----------------''
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40

''Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal cScan As Byte, _
''                                ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
                                ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

''----------------''
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, _
  ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, _
  lpReserved As Any) As Long
                                
Public Sub SHCopyFile(ByVal from_file As String, ByVal to_file As String)
Dim sh_op As SHFILEOPSTRUCT

    With sh_op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = from_file & vbNullChar & vbNullChar
        .pTo = to_file & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With

    SHFileOperation sh_op
End Sub

''----------------''
''----------------''
Public Sub CopyFileOBJ(fileSRC As String, fileDIS As String)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(fileSRC)
    s = f.Copy(fileDIS, True)
    ''MsgBox s
    
End Sub


''----------------''
''----------------''

Function DeleteDirFiles(ByVal dir_name As String)
Dim File_name As String
Dim files As Collection
Dim i As Integer

    ' Get a list of files it contains.
    Set files = New Collection
    File_name = dir_name
    Do While Len(File_name) > 0
        If (File_name <> "..") And (File_name <> ".") Then
            files.Add dir_name & "\" & File_name
        End If
        File_name = Dir$()
    Loop

    ' Delete the files.
    For i = 1 To files.Count
        File_name = files(i)
        ' See if it is a directory.
        If GetAttr(File_name) And vbDirectory Then
            ' It is a directory. Delete it.
            DeleteDirFiles File_name
        Else
            ' It's a file. Delete it.
'            lblStatus.Caption = file_name
'            lblStatus.Refresh
            SetAttr File_name, vbNormal
            Kill File_name
        End If
    Next i
    
    ''FileChk = False
    
    Set files = Nothing
    ' The directory is now empty. Delete it.
'    lblStatus.Caption = dir_name
'    lblStatus.Refresh
    'RmDir dir_name '''============================'''
End Function

''----------------''
''----------------''

Function FileExists(ByVal strPathName As String) As Boolean
  Dim af As Long
    af = GetFileAttributes(strPathName)
    FileExists = ((af <> -1) And af <> vbDirectory)
End Function

''----------------''
''----------------''

Function WaitForProcess(ByVal idProc As Long, Optional ByVal Sleep As Boolean) As Long
  
  Dim iExitCode As Long, hProc As Long
  Dim iResult As Long

    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, idProc)  ' Get process handle

    If Sleep Then  ' Sleep until process finishes
        
        iResult = WaitForSingleObject(hProc, INFINITE)
        If iResult = WAIT_FAILED Then Err.Raise Err.LastDllError

        GetExitCodeProcess hProc, iExitCode  ' Save the return code
    Else
        GetExitCodeProcess hProc, iExitCode  ' Get the return code
        
        Do While iExitCode = STILL_ACTIVE  ' Wait for process but relinquish time slice
            DoEvents
            GetExitCodeProcess hProc, iExitCode
            
            ''''''''''''''''''''''''''''''''''''''''{Debug!! Only!}
'            DoEvents
'            frmMain.txtData = frmMain.txtData + "W"
'            DoEvents
'            mSleep (200)
'            DoEvents
            ''''''''''''''''''''''''''''''''''''''''
        Loop
    End If

    CloseHandle hProc
    WaitForProcess = iExitCode    ' Return exit code
End Function

''----------------''
Function Check_Process(ByVal idProc As Long, Optional ByVal Sleep As Boolean) As Long
  
  Dim iExitCode As Long, hProc As Long
  Dim iResult As Long

    hProc = OpenProcess(PROCESS_ALL_ACCESS, False, idProc)  ' Get process handle

    If Sleep Then  ' Sleep until process finishes
        
        iResult = WaitForSingleObject(hProc, INFINITE)
        If iResult = WAIT_FAILED Then Err.Raise Err.LastDllError

        GetExitCodeProcess hProc, iExitCode  ' Save the return code
    Else
        GetExitCodeProcess hProc, iExitCode  ' Get the return code
        
    End If

    CloseHandle hProc
    
    Check_Process = iExitCode    ' Return exit code
End Function



'<<<<<<<<<<<<<Closing a DOS prompt window>>>>>>>>>>>
'
'After you run a DOS application in Windows 95, the MS-DOS Prompt window doesn't close.
' To prevent this behavior,
'  you can use the API to find the Window handle for the DOS prompt window,
' wait for the program to finish running, then zap the DOS prompt window into oblivion.
'
'This technique doesn't require any forms.
' It's just a simple VB 4.0 DLL with two properties:
'   the EXE name of the DOS program and the text that will appear
'   as the caption of the DOS prompt window displaying this application.
' The core of this app lies in three API calls. Place the following code in a standard module:
'
'Code
'
''Public Const WM_CLOSE = &H10


''The rest of the code goes in the following class module, named Cclose:
''
''Private m_sEXEName As String
''Private m_sDosCaption As String

Public Sub RunDosApp()

Dim vReturnValue As Variant
Dim lRet As Long
Dim i As Integer

    vReturnValue = Shell(m_sEXEName, 1) 'Run EXE
    AppActivate vReturnValue 'Activate EXE Window

    Do

        Sleep (1000)  ''10000

        lRet = FindWindow(vbNullString, m_sDosCaption)

        If (lRet <> 0) Then
            vReturnValue = SendMessage(lRet, WM_CLOSE, &O0, &O0)
            Exit Do  '==>
        End If

    Loop

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub KillDosApp(vRetVal As Variant)  ''LJS...

Dim vReturnValue As Variant
Dim lRet As Long
Dim i As Integer

    '--------------------------------------------NotUse??
    'AppActivate vRetVal 'Activate EXE Window

    Do

        Sleep (1000)  ''10000

        lRet = FindWindow(vbNullString, "M9S3V") ''m_sDosCaption)

        If (lRet <> 0) Then
            vReturnValue = SendMessage(lRet, WM_CLOSE, &O0, &O0)
            
            ''<<<<<<< Do ADD for Multiple-BUG >>>>>>>>>
            ''
            ''
            
            '' vReturnValue = SendMessage(lRet, WM_KEYDOWN, vbKeyEscape, &O0)
            
            Exit Do  '==>
        End If

    Loop

End Sub


Public Sub KillAppl(appl As String)   ''LJS...Test-Only...

Dim vReturnValue As Variant
Dim lRet As Long
Dim i As Integer

    Do
        Sleep (1000)  ''10000

        lRet = FindWindow(vbNullString, appl) ''m_sDosCaption)

        If (lRet <> 0) Then
            vReturnValue = SendMessage(lRet, WM_CLOSE, &O0, &O0)
            
            ''<<<<<<< Do ADD for Multiple-BUG >>>>>>>>>
            ''
            ''
            
            Exit Do  '==>
        End If

    Loop

End Sub

'<<<<<<<<<<<<<Closing a DOS prompt window>>>>>>>>>>>END


''==============================================================





Public Sub Log_RVSIBK(Log_data1 As String, Filename As String)
'======================================================================
    '로고파일 저장한다..
'======================================================================
    
  Dim f1
  Dim f2
  Dim Fname     As String
  Dim LogName   As String
  Dim Fint      As Integer
  Dim str1      As String
  Dim str2      As String
  

''DO: On-Err ???


        ''str2 = Format(Now, "YYYYMMDD hhmmss")
        ''str2 = Dac_date
        
        str1 = Left(str2, 8)

        Fname = App.Path
        Fint = Len(Filename)
        LogName = Mid$(Filename, 1, Fint - 4)
        Fname = Fname & "\" & LogName & ".log"

        If Not FileExists(Fname) Then
            f1 = FreeFile
            Open Fname For Binary Access Write As #f1
                ''Put #f1, , "DAC-LOG :: " + Fname + vbCrLf + vbCrLf
                Put #f1, , Log_data1$
            Close #f1
            DoEvents
            'Sleep 10
        Else
    
            f2 = FreeFile
            Open Fname For Binary Access Write As #f2
                Seek #f2, LOF(f2) + 1
                Put #f2, , vbCrLf & Log_data1$
            Close #f2
            DoEvents
        
        End If

        'Sleep 10
        
End Sub







'==============================================================







