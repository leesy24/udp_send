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
   StartUpPosition =   3  'Windows �⺻��
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
      Style           =   1  '�׷���
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
      ScrollBars      =   3  '�����
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
      Align           =   1  '�� ����
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
            Key             =   "�� ����"
            Object.ToolTipText     =   "�� ����"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�μ�"
            Object.ToolTipText     =   "�μ�"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�߶󳻱�"
            Object.ToolTipText     =   "�߶󳻱�"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "�ٿ��ֱ�"
            Object.ToolTipText     =   "�ٿ��ֱ�"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����Ӳ�"
            Object.ToolTipText     =   "����Ӳ�"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "����"
            Object.ToolTipText     =   "����"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "���� ����"
            Object.ToolTipText     =   "���� ����"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "��� ����"
            Object.ToolTipText     =   "��� ����"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "������ ����"
            Object.ToolTipText     =   "������ ����"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  '�Ʒ� ����
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
            Text            =   "����"
            TextSave        =   "����"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "2009-08-06"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "���� 4:33"
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
      Caption         =   "����(&F)"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "����(&O)..."
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "�ݱ�(&C)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "�ٸ� �̸����� ����(&A)..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "��� ����(&L)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "�Ӽ�(&I)"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "������ ����(&U)"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "�μ� �̸�����(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "�μ�(&P)..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "������(&D)..."
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
         Caption         =   "������(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "����(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "���� ����(&T)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "���� ǥ����(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "���� ��ħ(&R)"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "�ɼ�(&O)..."
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "�� ������(&W)"
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
        Case "�� ����"
            '�۾�: '�� ����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'�� ����' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "����"
            mnuFileOpen_Click
        Case "����"
            mnuFileSave_Click
        Case "�μ�"
            mnuFilePrint_Click
        Case "�߶󳻱�"
            '�۾�: '�߶󳻱�' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'�߶󳻱�' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "����"
            '�۾�: '����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'����' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "�ٿ��ֱ�"
            '�۾�: '�ٿ��ֱ�' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'�ٿ��ֱ�' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "����"
            '�۾�: '����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'����' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "����Ӳ�"
            '�۾�: '����Ӳ�' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'����Ӳ�' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "����"
            '�۾�: '����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'����' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "���� ����"
            '�۾�: '���� ����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'���� ����' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "��� ����"
            '�۾�: '��� ����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'��� ����' ���� �ڵ带 �߰��Ͻʽÿ�."
        Case "������ ����"
            '�۾�: '������ ����' ���� �ڵ带 �߰��Ͻʽÿ�.
            MsgBox "'������ ����' ���� �ڵ带 �߰��Ͻʽÿ�."
    End Select
End Sub

Private Sub mnuViewWebBrowser_Click()
    '�۾�: 'mnuViewWebBrowser_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewWebBrowser_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewOptions_Click()
    '�۾�: 'mnuViewOptions_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewOptions_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuViewRefresh_Click()
    '�۾�: 'mnuViewRefresh_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuViewRefresh_Click' �ڵ带 �߰��Ͻʽÿ�."
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
    '���� ��ε��մϴ�.
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    '�۾�: 'mnuFileSend_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSend_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFilePrint_Click()
    '�۾�: 'mnuFilePrint_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFilePrint_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFilePrintPreview_Click()
    '�۾�: 'mnuFilePrintPreview_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFilePrintPreview_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "������ ����"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    '�۾�: 'mnuFileProperties_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileProperties_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileSaveAll_Click()
    '�۾�: 'mnuFileSaveAll_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSaveAll_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileSaveAs_Click()
    '�۾�: 'mnuFileSaveAs_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSaveAs_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileSave_Click()
    '�۾�: 'mnuFileSave_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileSave_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileClose_Click()
    '�۾�: 'mnuFileClose_Click' �ڵ带 �߰��Ͻʽÿ�.
    MsgBox "'mnuFileClose_Click' �ڵ带 �߰��Ͻʽÿ�."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    With dlgCommonDialog
        .DialogTitle = "����"
        .CancelError = False
        '�۾�: Common Dialog ��Ʈ���� �÷��׿� Ư���� �����մϴ�.
        .Filter = "��� ����(*.*)|*.*"
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
    '�۾�: �ڵ带 �߰��Ͽ� ���� �ִ� ������ ó���մϴ�.
            
            
            
            
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

