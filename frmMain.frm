VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Windows�����Զ����������V2.1 ��ʾ"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   9240
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Text            =   "G:\Program Files (x86)\Tencent\QQ\Bin\qq.exe"
      Top             =   360
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "111222"
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "1324567"
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ��½"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "·��:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   6
      Top             =   435
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   1395
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "QQ:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   1
      Top             =   915
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim w As New clsWindow

Private Sub Command1_Click()
    loginQQ Text1.Text, Text2.Text
End Sub
'��½qq�����˺�����ĺ���
Private Sub loginQQ(ByVal strName$, ByVal strMsg$)
    Shell Text3.Text, 1
    Do While w.Width <> 495
        w.GetWindowByTitle "QQ", 1
    Loop
    w.Focus
    w.ClickPoint 163, 242, , , 200, 200
    SendKeys "{DEL}"
    SendKeys Text1.Text
    SendKeys "{TAB}"
    SendKeys Text2.Text
    SendKeys "{ENTER}"
End Sub

Private Sub Form_Load()
    Me.Show
    w.Load(Me.hWnd).FadeIn  '����
End Sub
