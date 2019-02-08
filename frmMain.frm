VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Project1"
   ClientHeight    =   5910
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtLport 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Text            =   "4321"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdUDPopen 
      Caption         =   "UDPopen"
      Height          =   375
      Left            =   3840
      MaskColor       =   &H8000000A&
      Style           =   1  '그래픽
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "SEND"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtTime 
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Text            =   "200"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtRport 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Text            =   "4101"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtRip 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Text            =   "192.168.0.255"
      Top             =   720
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7680
      Top             =   480
   End
   Begin VB.TextBox Text2 
      Height          =   3495
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  '양방향
      TabIndex        =   3
      Top             =   1920
      Width           =   8655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   7455
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  '위 맞춤
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "새 파일"
            Object.ToolTipText     =   "새 파일"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "열기"
            Object.ToolTipText     =   "열기"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "저장"
            Object.ToolTipText     =   "저장"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "인쇄"
            Object.ToolTipText     =   "인쇄"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "잘라내기"
            Object.ToolTipText     =   "잘라내기"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "복사"
            Object.ToolTipText     =   "복사"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "붙여넣기"
            Object.ToolTipText     =   "붙여넣기"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "굵게"
            Object.ToolTipText     =   "굵게"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "기울임꼴"
            Object.ToolTipText     =   "기울임꼴"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "밑줄"
            Object.ToolTipText     =   "밑줄"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "왼쪽 맞춤"
            Object.ToolTipText     =   "왼쪽 맞춤"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "가운데 맞춤"
            Object.ToolTipText     =   "가운데 맞춤"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "오른쪽 맞춤"
            Object.ToolTipText     =   "오른쪽 맞춤"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '아래 맞춤
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5640
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10636
            Text            =   "상태"
            TextSave        =   "상태"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2009-08-06"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "오후 4:33"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   8160
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   8640
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0112
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0224
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0336
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0448
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":055A
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066C
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0890
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A2
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AB4
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BC6
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CD8
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "열기(&O)..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "닫기(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "저장(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "다른 이름으로 저장(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "모두 저장(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "속성(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "페이지 설정(&U)"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "인쇄 미리보기(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "인쇄(&P)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "보내기(&D)..."
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "끝내기(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "보기(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "도구 모음(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "상태 표시줄(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "새로 고침(&R)"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "옵션(&O)..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "웹 브라우저(&W)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


''
''


Option Explicit

''


Dim f1
    
Dim VarString As String


Private Sub cmdSend_Click()


    Timer1.Interval = txtTime.Text

    If Timer1.Enabled = True Then
        Timer1.Enabled = False
        cmdSend.BackColor = &H8000000F
    Else
        Timer1.Enabled = True
        cmdSend.BackColor = vbBlue
    End If


End Sub

Private Sub cmdUDPopen_Click()


        With Winsock1
            .RemoteHost = Trim$(txtRip)
            .RemotePort = Trim$(txtRport)
            .LocalPort = Trim$(txtLport)

            .Bind .LocalPort
            
            cmdUDPopen.BackColor = vbBlue
            
        End With

End Sub

Private Sub Form_Load()
    
''    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
''    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
''    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
''    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    
    Text1.Text = ""
    


    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer


                
    Close #f1
    '''''''''


    'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "새 파일"
            '작업: '새 파일' 단추 코드를 추가하십시오.
            MsgBox "'새 파일' 단추 코드를 추가하십시오."
        Case "열기"
            mnuFileOpen_Click
        Case "저장"
            mnuFileSave_Click
        Case "인쇄"
            mnuFilePrint_Click
        Case "잘라내기"
            '작업: '잘라내기' 단추 코드를 추가하십시오.
            MsgBox "'잘라내기' 단추 코드를 추가하십시오."
        Case "복사"
            '작업: '복사' 단추 코드를 추가하십시오.
            MsgBox "'복사' 단추 코드를 추가하십시오."
        Case "붙여넣기"
            '작업: '붙여넣기' 단추 코드를 추가하십시오.
            MsgBox "'붙여넣기' 단추 코드를 추가하십시오."
        Case "굵게"
            '작업: '굵게' 단추 코드를 추가하십시오.
            MsgBox "'굵게' 단추 코드를 추가하십시오."
        Case "기울임꼴"
            '작업: '기울임꼴' 단추 코드를 추가하십시오.
            MsgBox "'기울임꼴' 단추 코드를 추가하십시오."
        Case "밑줄"
            '작업: '밑줄' 단추 코드를 추가하십시오.
            MsgBox "'밑줄' 단추 코드를 추가하십시오."
        Case "왼쪽 맞춤"
            '작업: '왼쪽 맞춤' 단추 코드를 추가하십시오.
            MsgBox "'왼쪽 맞춤' 단추 코드를 추가하십시오."
        Case "가운데 맞춤"
            '작업: '가운데 맞춤' 단추 코드를 추가하십시오.
            MsgBox "'가운데 맞춤' 단추 코드를 추가하십시오."
        Case "오른쪽 맞춤"
            '작업: '오른쪽 맞춤' 단추 코드를 추가하십시오.
            MsgBox "'오른쪽 맞춤' 단추 코드를 추가하십시오."
    End Select
End Sub

Private Sub mnuViewWebBrowser_Click()
    '작업: 'mnuViewWebBrowser_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewWebBrowser_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewOptions_Click()
    '작업: 'mnuViewOptions_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewOptions_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewRefresh_Click()
    '작업: 'mnuViewRefresh_Click' 코드를 추가하십시오.
    MsgBox "'mnuViewRefresh_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    '폼을 언로드합니다.
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    '작업: 'mnuFileSend_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSend_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFilePrint_Click()
    '작업: 'mnuFilePrint_Click' 코드를 추가하십시오.
    MsgBox "'mnuFilePrint_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFilePrintPreview_Click()
    '작업: 'mnuFilePrintPreview_Click' 코드를 추가하십시오.
    MsgBox "'mnuFilePrintPreview_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "페이지 설정"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    '작업: 'mnuFileProperties_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileProperties_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileSaveAll_Click()
    '작업: 'mnuFileSaveAll_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSaveAll_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileSaveAs_Click()
    '작업: 'mnuFileSaveAs_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSaveAs_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileSave_Click()
    '작업: 'mnuFileSave_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileSave_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileClose_Click()
    '작업: 'mnuFileClose_Click' 코드를 추가하십시오.
    MsgBox "'mnuFileClose_Click' 코드를 추가하십시오."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "열기"
        .CancelError = False
        '작업: Common Dialog 컨트롤의 플래그와 특성을 설정합니다.
        .Filter = "모든 파일(*.*)|*.*"
        .ShowOpen
        If Len(.Filename) = 0 Then
            Exit Sub
        End If
        sFile = .Filename
    
        Text1.Text = .Filename
        ''''''''''''''''''''''
    
    
        If FileExists(.Filename) Then
             
             f1 = FreeFile
            Open .Filename For Input As #f1
        
        End If
        '''
    
    
    
    End With
    '작업: 코드를 추가하여 열려 있는 파일을 처리합니다.
            
            
            
            
End Sub








Private Sub Timer1_Timer()


        If Not EOF(f1) Then
            
            Line Input #1, VarString
            ''Debug.Print VarString

            If Len(VarString) > 0 Then
            
                Winsock1.SendData VarString & vbCrLf
                ''''''''''''''''''''''''''''''''''''
                DoEvents
                
''                Text2.Text = Text2.Text & VarString & vbCrLf
''                Text2.SelStart = Len(Text2.Text)
                Text2.Text = VarString & vbCrLf

                
                DoEvents
                
            End If
            
        End If




End Sub

