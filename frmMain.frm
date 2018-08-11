VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0080FFFF&
   Caption         =   "English Test"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "VNI-Times"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtInstruction 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   7560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   7695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Th&oa�t"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   9360
      Width           =   7095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "T&r�� giu�p"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   8280
      Width           =   7095
   End
   Begin VB.CommandButton cmdDic 
      Caption         =   "T�� �&ie�n"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Width           =   7095
   End
   Begin VB.CommandButton cmdAdj 
      Caption         =   "T�n&h t�� tra�i ngh�a"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   6120
      Width           =   7095
   End
   Begin VB.CommandButton cmdIrr 
      Caption         =   "�o�&ng t�� ba�t quy ta�c"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   7095
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "&Gi��i t�� th�ch h��p"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   7095
   End
   Begin VB.CommandButton cmdEV 
      Caption         =   "Tie�ng &Anh sang tie�ng Vie�t"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   7095
   End
   Begin VB.CommandButton cmdVE 
      Caption         =   "&Tie�ng Vie�t sang tie�ng Anh"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "ENGLISH TEST"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public frm As Byte
Private Sub cmdAdj_Click()
    Unload frmMain
    Load frmAdj
End Sub

Private Sub cmdAdj_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "T�m t�nh t�� tra�i ngh�a v��i t�nh t�� �a� cho."
End Sub

Private Sub cmdDic_Click()
    Unload frmMain
    Load frmDic
End Sub

Private Sub cmdDic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "Xem t�� �ie�n cu�a English Test. Ba�n cu�ng co� the� the�m, thay �o�i hoa�c xo�a t��."
End Sub

Private Sub cmdEV_Click()
    Unload frmMain
    Load frmEV
End Sub

Private Sub cmdEV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "T�m t�� tie�ng Vie�t co� ngh�a th�ch h��p v��i t�� tie�ng Anh."
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "Thoa�t English Test."
End Sub

Private Sub cmdHelp_Click()
    Load frmHelp
End Sub

Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "H���ng da�n s�� du�ng ch��ng tr�nh English Test."
End Sub

Private Sub cmdIrr_Click()
    Unload frmMain
    Load frmIrr
End Sub

Private Sub cmdIrr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "Ba�n co� the� nh�� ����c bao nhie�u �o�ng t�� ba�t quy ta�c?"
End Sub

Private Sub cmdPre_Click()
    Unload frmMain
    Load frmPre
End Sub

Private Sub cmdPre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "T�m gi��i t�� th�ch h��p v��i danh t��, �o�ng t�� va� t�nh t��."
End Sub

Private Sub cmdVE_Click()
    Unload frmMain
    Load frmVE
End Sub

Private Sub cmdVE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "T�m t�� tie�ng Anh co� ngh�a th�ch h��p v��i t�� tie�ng Vie�t."
End Sub

Private Sub Form_Load()
    frmMain.Visible = True
    txtInstruction = "Welcome to English Test"
    frm = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInstruction = "Welcome to English Test"
End Sub
