VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ȭ�� ���
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.PictureBox picLogo 
         Height          =   2385
         Left            =   510
         Picture         =   "frmSplash.frx":0000
         ScaleHeight     =   2325
         ScaleWidth      =   1755
         TabIndex        =   2
         Top             =   855
         Width           =   1815
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  '������ ����
         Caption         =   "�� ��ǰ�� ���� ����ڿ��� ����� �㰡�Ǿ����ϴ�."
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   1
         Tag             =   "�� ��ǰ�� ���� ����ڿ��� ����� �㰡�Ǿ����ϴ�."
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   32.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2670
         TabIndex        =   9
         Tag             =   "��ǰ"
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "����"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2505
         TabIndex        =   8
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   2880
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "�÷���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6015
         TabIndex        =   7
         Tag             =   "�÷���"
         Top             =   2400
         Width           =   990
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  '������ ����
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6495
         TabIndex        =   6
         Tag             =   "����"
         Top             =   2760
         Width           =   510
      End
      Begin VB.Label lblWarning 
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Tag             =   "���"
         Top             =   3720
         Width           =   6855
      End
      Begin VB.Label lblCompany 
         Caption         =   "ȸ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   5
         Tag             =   "ȸ��"
         Top             =   3330
         Width           =   2415
      End
      Begin VB.Label lblCopyright 
         Caption         =   "���۱�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4710
         TabIndex        =   4
         Tag             =   "���۱�"
         Top             =   3120
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

