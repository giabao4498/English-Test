VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H0080FFFF&
   Caption         =   "Tro giup"
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
   Begin TabDlg.SSTab tabHelp 
      Height          =   10455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   18441
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   750
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Giôùi thieäu"
      TabPicture(0)   =   "frmHelp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtIntro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vieät - Anh"
      TabPicture(1)   =   "frmHelp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtVE"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Anh - Vieät"
      TabPicture(2)   =   "frmHelp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Giôùi töø"
      TabPicture(3)   =   "frmHelp.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Irregular verb"
      TabPicture(4)   =   "frmHelp.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Tính töø traùi"
      TabPicture(5)   =   "frmHelp.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Töø ñieån"
      TabPicture(6)   =   "frmHelp.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      Begin VB.TextBox txtVE 
         Height          =   9735
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   14895
      End
      Begin VB.TextBox txtIntro 
         Height          =   9735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmHelp.frx":00C4
         Top             =   600
         Width           =   14895
      End
   End
   Begin VB.Menu mnuBack 
      Caption         =   "&Tro vê"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "T&hoát"
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmHelp.Visible = True
End Sub

Private Sub mnuBack_Click()
    Unload frmHelp
    Select Case frmMain.frm
        Case 1
            If frmVE.lblTime.Caption <> "" Then
                frmVE.tmrTime.Enabled = True
            End If
        Case 2
            If frmEV.lblTime.Caption <> "" Then
                frmEV.tmrTime.Enabled = True
            End If
        Case 3
            If frmPre.lblTime.Caption <> "" Then
                frmPre.tmrTime.Enabled = True
            End If
    End Select
End Sub

Private Sub mnuExit_Click()
    Dim Button As Byte
    Button = MsgBox("Ban có chac chan muôn thoát không?", vbYesNo, "Thoát")
    If Button = 6 Then
        End
    End If
End Sub
