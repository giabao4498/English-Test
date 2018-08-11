VERSION 5.00
Begin VB.Form frmEV 
   BackColor       =   &H0080FFFF&
   Caption         =   "Tiêng Anh sang tiêng Viêt"
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
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   5880
   End
   Begin VB.CommandButton cmdAns 
      Enabled         =   0   'False
      Height          =   855
      Index           =   4
      Left            =   11280
      TabIndex        =   13
      Top             =   9120
      Width           =   3975
   End
   Begin VB.CommandButton cmdAns 
      Enabled         =   0   'False
      Height          =   855
      Index           =   3
      Left            =   7200
      TabIndex        =   12
      Top             =   9120
      Width           =   3975
   End
   Begin VB.CommandButton cmdAns 
      Enabled         =   0   'False
      Height          =   855
      Index           =   2
      Left            =   11280
      TabIndex        =   11
      Top             =   7920
      Width           =   3975
   End
   Begin VB.CommandButton cmdAns 
      Enabled         =   0   'False
      Height          =   855
      Index           =   1
      Left            =   7200
      TabIndex        =   10
      Top             =   7920
      Width           =   3975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Baét ñaàu"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13440
      Top             =   0
   End
   Begin VB.ListBox lstResults 
      Height          =   8040
      ItemData        =   "frmEV.frx":0000
      Left            =   120
      List            =   "frmEV.frx":0002
      TabIndex        =   7
      ToolTipText     =   "Danh sách các tu tra loi sai hoac không tra loi"
      Top             =   2280
      Width           =   7095
   End
   Begin VB.TextBox txtAmount 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      Caption         =   "Nhöõng töø khoâng traû lôøi ñöôïc"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label lblIncorrect 
      Caption         =   "Sai: "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   11640
      TabIndex        =   6
      ToolTipText     =   "Sô tu tra loi sai"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Label lblCorrect 
      Caption         =   "Ñuùng: "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   7560
      TabIndex        =   5
      ToolTipText     =   "Sô tu tra loi dung"
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label lblAns 
      Caption         =   "Ñaõ traû lôøi: "
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   4
      ToolTipText     =   "Sô tu dã tra loi"
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Image imgPrevious 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   9960
      Picture         =   "frmEV.frx":0004
      ToolTipText     =   "Tu phía truoc"
      Top             =   5160
      Width           =   1260
   End
   Begin VB.Image imgNext 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   11280
      Picture         =   "frmEV.frx":4432
      ToolTipText     =   "Tu tiêp theo"
      Top             =   5160
      Width           =   1260
   End
   Begin VB.Label lblWord 
      Alignment       =   2  'Center
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
      Left            =   7200
      TabIndex        =   3
      Top             =   6360
      Width           =   8055
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13920
      TabIndex        =   2
      ToolTipText     =   "Thoi gian còn lai"
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblAmount 
      Alignment       =   2  'Center
      Caption         =   "Nhaäp soá löôïng töø:"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Tro vê bang chon chính"
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Kiêm tra"
      Begin VB.Menu mnuTestVE 
         Caption         =   "&Tiêng Viêt sang tiêng Anh"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuTestPre 
         Caption         =   "&Gioi tu thích hop"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuTestIrr 
         Caption         =   "&Dông tu bât quy tac"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuTestAdj 
         Caption         =   "Tí&nh tu trái nghia"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuDic 
      Caption         =   "T&u diên"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "T&ro giúp"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "T&hoát"
   End
End
Attribute VB_Name = "frmEV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type record
    e As String
    e2 As String
    e3 As String
    v As String
    e4 As String
    v2 As String
    v2_ As String
    v3 As String
    v3_ As String
    pre As String
    pre2 As String
    chosen As Boolean
    ans(1 To 4) As String
    main As String
End Type
Dim n As Integer, a() As record, m As Integer, b() As record, tt As Integer, t As Integer, f As Integer, m_ As Integer, sec As Integer, min As Byte
Sub inc()
    Do
        If tt < m - 1 Then
            tt = tt + 1
        Else
            tt = 0
        End If
    Loop Until b(tt).chosen = True
End Sub
Sub dec()
    Do
        If tt > 0 Then
            tt = tt - 1
        Else
            tt = m - 1
        End If
    Loop Until b(tt).chosen = True
End Sub
Sub over(nd As String, title As String)
    tmrTime.Enabled = False
    MsgBox nd, , title
    mnuDic.Enabled = True
    cmdStart.Enabled = True
    lblTime.Caption = ""
    lblAmount.Enabled = True
    txtAmount.Enabled = True
    lblWord.Caption = ""
    Dim i As Byte
    For i = 1 To 4
        With cmdAns(i)
            .Enabled = False
            .Caption = ""
        End With
    Next i
End Sub
Function kt(i As Integer, j As Integer) As Boolean
    Dim k As Byte
    With b(i)
        For k = 1 To 4
            If (k <> j) And (.ans(k) = .ans(j)) Then
                kt = False
                Exit Function
            End If
        Next k
    End With
    kt = True
End Function
Sub init()
    lblWord.Caption = b(tt).main
    If b(tt).v2 <> "" Then
        With lblWord
            .Caption = .Caption & " - " & b(tt).v2 & " - " & b(tt).v3
        End With
    End If
    Dim i As Byte
    For i = 1 To 4
        cmdAns(i).Caption = b(tt).ans(i)
    Next i
End Sub
Private Sub cmdAns_Click(Index As Integer)
    tmrTime.Enabled = False
    Dim i As Integer
    m_ = m_ + 1
    b(tt).chosen = False
    With lblAns
        .Caption = Mid(.Caption, 1, 15) & m_ & "/" & m
    End With
    If cmdAns(Index).Caption = b(tt).v Then
        t = t + 1
        cmdAns(Index).Caption = "Ñuùng"
        With lblCorrect
            .Caption = Left(.Caption, 7) & t
        End With
    Else
        f = f + 1
        For i = 1 To 4
            With cmdAns(i)
                If .Caption <> b(tt).v Then
                    .Enabled = False
                End If
            End With
        Next i
        With lblIncorrect
            .Caption = Left(.Caption, 5) & f
        End With
        With b(tt)
            lstResults.AddItem .v & "     " & .e
        End With
    End If
    tmrWait.Enabled = True
End Sub

Private Sub cmdAns_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If tmrTime.Enabled = True Then
        Select Case KeyCode
            Case 188
                Replace txtAns, ",", ""
                Replace txtAns, "<", ""
                imgPrevious_Click
            Case 190
                Replace txtAns, ".", ""
                Replace txtAns, ">", ""
                imgNext_Click
        End Select
    End If
End Sub

Private Sub cmdStart_Click()
    If txtAmount = "" Then
        MsgBox "Ban chua nhâp sô luong tu!", vbExclamation, "Sô luong tu không hop lê!"
        m_ = -1
        Exit Sub
    End If
    Dim i As Integer
    For i = 1 To Len(txtAmount)
        If Mid(txtAmount, i, 1) Like "[!0-9]" Then
            MsgBox "Sô luong tu không hop lê! Vui lòng nhâp lai!", vbExclamation, "Sô luong tu không hop lê!"
            txtAmount = ""
            m_ = -1
            Exit Sub
        End If
    Next i
    mnuDic.Enabled = False
    lblAmount.Enabled = False
    txtAmount.Enabled = False
    cmdStart.Enabled = False
    lblAns.Enabled = True
    lblResults.Enabled = True
    lstResults.Clear
    With lblCorrect
        .Enabled = True
        .Caption = "Ñuùng: 0"
    End With
    With lblIncorrect
        .Enabled = True
        .Caption = "Sai: 0"
    End With
    imgPrevious.Enabled = True
    imgNext.Enabled = True
    For i = 1 To 4
        cmdAns(i).Enabled = True
    Next i
    For i = 0 To n - 1
        a(i).chosen = False
    Next i
    m = Val(txtAmount)
    Dim j As Integer
    For i = 1 To m
        ReDim Preserve b(i) As record
        For j = 1 To 4
            b(i - 1).ans(j) = ""
        Next j
        Do
            j = Int(Rnd * n)
        Loop Until a(j).chosen = False
        a(j).chosen = True
        b(i - 1) = a(j)
        With b(i - 1)
            .ans(Int(Rnd * 4 + 1)) = .v
            Do
                j = Int(Rnd * 3)
                Select Case j
                    Case 0
                        .main = .e
                    Case 1
                        .main = .e2
                    Case 2
                        .main = .e3
                End Select
            Loop Until .main <> ""
        End With
        For j = 1 To 4
            If b(i - 1).ans(j) = "" Then
                Do
                    b(i - 1).ans(j) = a(Int(Rnd * n)).v
                Loop Until kt(i - 1, j)
            End If
        Next j
    Next i
    m_ = 0
    lblAns.Caption = "Ñaõ traû lôøi: 0/" & m
    t = 0
    f = 0
    tt = 0
    sec = 5 * m
    min = sec \ 60
    sec = sec Mod 60
    With lblTime
        .Caption = min & ":"
        If sec < 10 Then
            .Caption = .Caption & "0"
        End If
        .Caption = .Caption & sec
    End With
    init
    tmrTime.Enabled = True
End Sub
Private Sub Form_Load()
    frmEV.Visible = True
    Open "DIC.DAT" For Input As #1
    Lock #1
    Dim tg As String
    Line Input #1, tg
    n = Val(tg)
    Dim i As Integer
    For i = 1 To n
        ReDim Preserve a(i) As record
        With a(i - 1)
            Line Input #1, .e
            Line Input #1, .e2
            Line Input #1, .e3
            Line Input #1, .v
            Line Input #1, .e4
            Line Input #1, .v2
            Line Input #1, .v2_
            Line Input #1, .v3
            Line Input #1, .v3_
            Line Input #1, .pre
            Line Input #1, .pre2
        End With
    Next i
    Close #1
    Randomize
    frmMain.frm = 2
End Sub

Private Sub imgNext_Click()
    inc
    init
End Sub

Private Sub imgPrevious_Click()
    dec
    init
End Sub

Private Sub lstResults_KeyUp(KeyCode As Integer, Shift As Integer)
    If tmrTime.Enabled = True Then
        Select Case KeyCode
            Case 188
                Replace txtAns, ",", ""
                Replace txtAns, "<", ""
                imgPrevious_Click
            Case 190
                Replace txtAns, ".", ""
                Replace txtAns, ">", ""
                imgNext_Click
        End Select
    End If
End Sub

Private Sub mnuDic_Click()
    Unload frmEV
    Load frmDic
End Sub

Private Sub mnuExit_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban có chac chan muôn thoát không?", vbYesNo, "Thoát")
        If Button = 6 Then
            End
        Else
            tmrTime.Enabled = True
        End If
    Else
        End
    End If
End Sub

Private Sub mnuHelp_Click()
    With tmrTime
        If .Enabled = True Then
            .Enabled = False
        End If
    End With
    Load frmHelp
End Sub

Private Sub mnuMain_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban có chac chan muôn tro vê bang chon chính không?", vbYesNo, "Tro vê bang chon chính")
        If Button = 6 Then
            Unload frmEV
            Load frmMain
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmEV
        Load frmMain
    End If
End Sub

Private Sub mnuTestAdj_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban có chac chan muôn kiêm tra tính tu trái nghia không?", vbYesNo, "Kiêm tra tính tu trái nghia")
        If Button = 6 Then
            Unload frmEV
            Load frmAdj
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmEV
        Load frmAdj
    End If
End Sub
Private Sub mnuTestIrr_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban có chac chan muôn kiêm tra dông tu bât quy tac không?", vbYesNo, "Kiêm tra dông tu bât quy tac")
        If Button = 6 Then
            Unload frmEV
            Load frmIrr
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmEV
        Load frmIrr
    End If
End Sub

Private Sub mnuTestPre_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban có chac chan muôn kiêm tra gioi tu không?", vbYesNo, "Kiêm tra gioi tu")
        If Button = 6 Then
            Unload frmEV
            Load frmPre
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmEV
        Load frmPre
    End If
End Sub

Private Sub mnuTestVE_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban có chac chan muôn kiêm tra tiêng Viêt sang tiêng Anh không?", vbYesNo, "Kiêm tra tiêng Anh sang tiêng Viêt")
        If Button = 6 Then
            Unload frmEV
            Load frmVE
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmEV
        Load frmVE
    End If
End Sub

Private Sub tmrTime_Timer()
    If (min = 0) And (sec = 0) Then
        over "Ban dã hêt thoi gian!", "Hêt gio!"
        Dim i As Integer
        For i = 0 To m - 1
            If b(i).chosen Then
                lstResults.AddItem b(i).v & "     " & b(i).e
            End If
        Next i
        Exit Sub
    End If
    If sec > 0 Then
        sec = sec - 1
    Else
        min = min - 1
        sec = 59
    End If
    With lblTime
        .Caption = min & ":"
        If sec < 10 Then
            .Caption = .Caption & "0"
        End If
        .Caption = .Caption & sec
    End With
End Sub
Private Sub tmrWait_Timer()
    tmrWait.Enabled = False
    Dim i As Integer
    For i = 1 To 4
        With cmdAns(i)
            If .BackColor <> &H8000000F Then
                .BackColor = &H8000000F
            End If
        End With
    Next i
    For i = 1 To 4
        With cmdAns(i)
            .Caption = ""
            .Enabled = True
        End With
    Next i
    If m_ < m Then
        imgNext_Click
    Else
        over "Ban dã hoàn thành bài kiêm tra!", "Hoàn thành!"
        With lstResults
            If .ListCount = 0 Then
                .AddItem "Khoâng coù töø sai"
            End If
        End With
        Exit Sub
    End If
    tmrTime.Enabled = True
End Sub
Private Sub txtAmount_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13) And (m_ > -1) Then
        cmdStart_Click
    Else
        m_ = 0
    End If
End Sub
