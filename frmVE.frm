VERSION 5.00
Begin VB.Form frmVE 
   BackColor       =   &H0080FFFF&
   Caption         =   "Ti�ng Vi�t sang ti�ng Anh"
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
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13440
      Top             =   0
   End
   Begin VB.ListBox lstResults 
      Height          =   8040
      ItemData        =   "frmVE.frx":0000
      Left            =   120
      List            =   "frmVE.frx":0002
      TabIndex        =   10
      ToolTipText     =   "Danh s�ch c�c tu tra loi sai hoac kh�ng tra loi"
      Top             =   2280
      Width           =   7095
   End
   Begin VB.TextBox txtAns 
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
      Left            =   7920
      TabIndex        =   4
      Text            =   "Nha�n va�o �a�y �e� ba�t �a�u"
      Top             =   2760
      Width           =   5775
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
      Caption         =   "Nh��ng t�� kho�ng tra� l��i ����c"
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
      TabIndex        =   11
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label lblKey 
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
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   8640
      TabIndex        =   9
      ToolTipText     =   "D�p �n ch�nh x�c"
      Top             =   7920
      Width           =   5775
   End
   Begin VB.Label lblResult 
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
      Left            =   10200
      TabIndex        =   8
      ToolTipText     =   "K�t qua"
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Image imgEnter 
      Enabled         =   0   'False
      Height          =   870
      Left            =   13800
      Picture         =   "frmVE.frx":0004
      ToolTipText     =   "Tra loi"
      Top             =   2760
      Width           =   900
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
      TabIndex        =   7
      ToolTipText     =   "S� tu tra loi sai"
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label lblCorrect 
      Caption         =   "�u�ng: "
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
      TabIndex        =   6
      ToolTipText     =   "S� tu tra loi dung"
      Top             =   5280
      Width           =   3975
   End
   Begin VB.Label lblAns 
      Caption         =   "�a� tra� l��i: "
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
      Left            =   7200
      TabIndex        =   5
      ToolTipText     =   "S� tu d� tra loi"
      Top             =   3960
      Width           =   5535
   End
   Begin VB.Image imgPrevious 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   12720
      Picture         =   "frmVE.frx":290E
      ToolTipText     =   "Tu ph�a truoc"
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Image imgNext 
      Enabled         =   0   'False
      Height          =   1035
      Left            =   14040
      Picture         =   "frmVE.frx":6D3C
      ToolTipText     =   "Tu ti�p theo"
      Top             =   3960
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
      Top             =   1560
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
      ToolTipText     =   "Thoi gian c�n lai"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblAmount 
      Alignment       =   2  'Center
      Caption         =   "Nha�p so� l���ng t��:"
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
      Caption         =   "&Tro v� bang chon ch�nh"
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Ki�m tra"
      Begin VB.Menu mnuTestEV 
         Caption         =   "&Ti�ng Anh sang ti�ng Vi�t"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTestPre 
         Caption         =   "&Gioi tu th�ch hop"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuTestIrr 
         Caption         =   "&D�ng tu b�t quy tac"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuTestAdj 
         Caption         =   "T�&nh tu tr�i nghia"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuDic 
      Caption         =   "T&u di�n"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "T&ro gi�p"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "T&ho�t"
   End
End
Attribute VB_Name = "frmVE"
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
    lblTime.Caption = ""
    lblAmount.Enabled = True
    txtAmount.Enabled = True
    lblWord.Caption = ""
    txtAns = "Nha�n va�o �a�y �e� ba�t �a�u"
    imgEnter.Enabled = False
    lblResult.Caption = ""
    lblKey.Caption = ""
End Sub

Private Sub Form_Load()
    frmVE.Visible = True
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
    frmMain.frm = 1
End Sub

Private Sub imgEnter_Click()
    m_ = m_ + 1
    b(tt).chosen = False
    With lblAns
        .Caption = Mid(.Caption, 1, 15) & m_ & "/" & m
    End With
    If (txtAns <> "") And ((txtAns = b(tt).e) Or (txtAns = b(tt).e2) Or (txtAns = b(tt).e3)) Then
        With lblResult
            .ForeColor = &HFF0000
            .Caption = "�u�ng"
        End With
        lblKey.Caption = ""
        t = t + 1
        With lblCorrect
            .Caption = Mid(.Caption, 1, 7) & t
        End With
    Else
        With lblResult
            .ForeColor = &HFF&
            .Caption = "Sai"
        End With
        lblKey.Caption = b(tt).e
        f = f + 1
        With lblIncorrect
            .Caption = Mid(.Caption, 1, 5) & f
        End With
        With b(tt)
            lstResults.AddItem .v & "     " & .e
        End With
    End If
    If m_ < m Then
        inc
        lblWord.Caption = b(tt).v
        txtAns = ""
    Else
        over "Ban d� ho�n th�nh b�i ki�m tra!", "Ho�n th�nh!"
        With lstResults
            If .ListCount = 0 Then
                .AddItem "Kho�ng co� t�� sai"
            End If
        End With
        txtAns = "Nha�n va�o �a�y �e� ba�t �a�u"
    End If
End Sub

Private Sub imgNext_Click()
    inc
    txtAns = ""
    lblWord.Caption = b(tt).v
End Sub

Private Sub imgPrevious_Click()
    dec
    txtAns = ""
    lblWord.Caption = b(tt).v
End Sub

Private Sub mnuDic_Click()
    Unload frmVE
    Load frmDic
End Sub

Private Sub mnuExit_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban c� chac chan mu�n tho�t kh�ng?", vbYesNo, "Tho�t")
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
        Button = MsgBox("Ban c� chac chan mu�n tro v� bang chon ch�nh kh�ng?", vbYesNo, "Tro v� bang chon ch�nh")
        If Button = 6 Then
            Unload frmVE
            Load frmMain
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmVE
        Load frmMain
    End If
End Sub

Private Sub mnuTestAdj_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban c� chac chan mu�n ki�m tra t�nh tu tr�i nghia kh�ng?", vbYesNo, "Ki�m tra t�nh tu tr�i nghia")
        If Button = 6 Then
            Unload frmVE
            Load frmAdj
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmVE
        Load frmAdj
    End If
End Sub

Private Sub mnuTestEV_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban c� chac chan mu�n ki�m tra ti�ng Anh sang ti�ng Vi�t kh�ng?", vbYesNo, "Ki�m tra ti�ng Anh sang ti�ng Vi�t")
        If Button = 6 Then
            Unload frmVE
            Load frmEV
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmVE
        Load frmEV
    End If
End Sub

Private Sub mnuTestIrr_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban c� chac chan mu�n ki�m tra d�ng tu b�t quy tac kh�ng?", vbYesNo, "Ki�m tra d�ng tu b�t quy tac")
        If Button = 6 Then
            Unload frmVE
            Load frmIrr
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmVE
        Load frmIrr
    End If
End Sub

Private Sub mnuTestPre_Click()
    If tmrTime.Enabled = True Then
        tmrTime.Enabled = False
        Dim Button As Byte
        Button = MsgBox("Ban c� chac chan mu�n ki�m tra gioi tu kh�ng?", vbYesNo, "Ki�m tra gioi tu")
        If Button = 6 Then
            Unload frmVE
            Load frmPre
        Else
            tmrTime.Enabled = True
        End If
    Else
        Unload frmVE
        Load frmPre
    End If
End Sub

Private Sub tmrTime_Timer()
    If (min = 0) And (sec = 0) Then
        over "Ban d� h�t thoi gian!", "H�t gio!"
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
Private Sub txtAns_Click()
    If tmrTime.Enabled = True Then
        Exit Sub
    End If
    txtAmount = RTrim(LTrim(txtAmount))
    If txtAmount = "" Then
        MsgBox "Ban chua nh�p s� luong tu!", vbExclamation, "S� luong tu kh�ng hop l�!"
        Exit Sub
    End If
    Dim i As Integer
    For i = 1 To Len(txtAmount)
        If Mid(txtAmount, i, 1) Like "[!0-9]" Then
            MsgBox "S� luong tu kh�ng hop l�! Vui l�ng nh�p lai!", vbExclamation, "S� luong tu kh�ng hop l�!"
            txtAmount = ""
            Exit Sub
        End If
    Next i
    mnuDic.Enabled = False
    lblAmount.Enabled = False
    txtAmount.Enabled = False
    lblResults.Enabled = True
    lstResults.Clear
    imgEnter.Enabled = True
    lblAns.Enabled = True
    imgPrevious.Enabled = True
    imgNext.Enabled = True
    With lblCorrect
        .Enabled = True
        .Caption = "�u�ng: "
    End With
    With lblIncorrect
        .Enabled = True
        .Caption = "Sai: "
    End With
    For i = 0 To n - 1
        a(i).chosen = False
    Next i
    m = Val(txtAmount)
    Dim j As Integer
    For i = 1 To m
        ReDim Preserve b(i) As record
        Do
            j = Int(Rnd * n)
        Loop Until a(j).chosen = False
        b(i - 1) = a(j)
        b(i - 1).chosen = True
        a(j).chosen = True
    Next i
    m_ = 0
    lblAns.Caption = "�a� tra� l��i: 0/" & m
    t = 0
    lblCorrect.Caption = lblCorrect.Caption & "0"
    f = 0
    lblIncorrect.Caption = lblIncorrect.Caption & "0"
    tt = 0
    sec = 10 * m
    min = sec \ 60
    sec = sec Mod 60
    With lblTime
        .Caption = min & ":"
        If sec < 10 Then
            .Caption = .Caption & "0"
        End If
        .Caption = .Caption & sec
    End With
    txtAns = ""
    lblWord.Caption = b(0).v
    tmrTime.Enabled = True
End Sub

Private Sub txtAns_KeyUp(KeyCode As Integer, Shift As Integer)
    If tmrTime.Enabled = True Then
        Select Case KeyCode
            Case 13
                imgEnter_Click
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
