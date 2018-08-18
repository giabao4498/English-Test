VERSION 5.00
Begin VB.Form frmDic 
   BackColor       =   &H0080FFFF&
   Caption         =   "Tu diên"
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
   Icon            =   "frmDic.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFind 
      Caption         =   "Tìm töø"
      Height          =   615
      Left            =   5160
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtExplain 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   5160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmDic.frx":0442
      Top             =   1560
      Width           =   7095
   End
   Begin VB.ListBox lstWord 
      Height          =   8880
      ItemData        =   "frmDic.frx":0450
      Left            =   120
      List            =   "frmDic.frx":0452
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Xoùa töø"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12360
      TabIndex        =   4
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Theâm töø môùi"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12360
      TabIndex        =   3
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Chænh söûa töø"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   12360
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   1
      Top             =   855
      Width           =   4935
   End
   Begin VB.Label lblWord 
      Alignment       =   2  'Center
      Caption         =   "Nhaäp töø caàn tìm"
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
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Menu mnuBack 
      Caption         =   "&Tro vê"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "T&ro giup"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "T&hoát"
   End
End
Attribute VB_Name = "frmDic"
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
End Type
Dim a() As record, n As Long
Sub sort()
    Dim i As Long, j As Long, tg As record
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If LCase(a(i).e) > LCase(a(j).e) Then
                tg = a(i)
                a(i) = a(j)
                a(j) = tg
            End If
        Next j
    Next i
    Open "DIC.DAT" For Output As #1
    Lock #1
    Print #1, n
    For i = 0 To n - 1
        With a(i)
            Print #1, .e
            Print #1, .e2
            Print #1, .e3
            Print #1, .v
            Print #1, .e4
            Print #1, .v2
            Print #1, .v2_
            Print #1, .v3
            Print #1, .v3_
            Print #1, .pre
            Print #1, .pre2
        End With
    Next i
    Close #1
End Sub
Public Sub change()
    Dim i As Long
    For i = 0 To n - 1
        With a(i)
            If .e = frmSet.b Then
                .e = frmSet.txtE
                .v = frmSet.txtV
                .e2 = frmSet.txtE2
                .e3 = frmSet.txtE3
                .e4 = frmSet.txtE4
                .v2 = frmSet.txtV2
                .v2_ = frmSet.txtV2_
                .v3 = frmSet.txtV3
                .v3_ = frmSet.txtV3_
                .pre = frmSet.txtPre
                .pre2 = frmSet.txtPre2
                lstWord.RemoveItem i
                lstWord.AddItem .e
                Exit For
            End If
        End With
    Next i
    sort
    txtWord = ""
    txtExplain = "Giaûi thích"
End Sub
Public Sub add()
    n = n + 1
    ReDim Preserve a(n) As record
    With a(n - 1)
        .e = frmSet.txtE
        .v = frmSet.txtV
        .e2 = frmSet.txtE2
        .e3 = frmSet.txtE3
        .e4 = frmSet.txtE4
        .v2 = frmSet.txtV2
        .v2_ = frmSet.txtV2_
        .v3 = frmSet.txtV3
        .v3_ = frmSet.txtV3_
        .pre = frmSet.txtPre
        .pre2 = frmSet.txtPre2
        lstWord.AddItem .e
    End With
    sort
    txtWord = ""
    txtExplain = "Giaûi thích"
End Sub

Private Sub cmdAdd_Click()
    Load frmSet
    With frmSet
        .Caption = "Thêm tu moi"
        .lblTitle = "Theâm töø môùi"
        .cmdDone.Caption = "Theâm töø"
    End With
End Sub

Private Sub cmdChange_Click()
    If lstWord.ListIndex > -1 Then
        Load frmSet
        With frmSet
            .Caption = "Chinh sua tu"
            .lblTitle = "Chænh söûa"
            .cmdDone.Caption = "Chænh söûa"
            .txtE = a(lstWord.ListIndex).e
            .txtV = a(lstWord.ListIndex).v
            .txtE2 = a(lstWord.ListIndex).e2
            .txtE3 = a(lstWord.ListIndex).e3
            .txtE4 = a(lstWord.ListIndex).e4
            .txtV2 = a(lstWord.ListIndex).v2
            .txtV2_ = a(lstWord.ListIndex).v2_
            .txtV3 = a(lstWord.ListIndex).v3
            .txtV3_ = a(lstWord.ListIndex).v3_
            .txtPre = a(lstWord.ListIndex).pre
            .txtPre2 = a(lstWord.ListIndex).pre2
            .b = .txtE
        End With
    Else
        MsgBox "Ban chua chon tu!", vbExclamation, "Chua chon tu!"
    End If
End Sub

Private Sub cmdDel_Click()
    If lstWord.ListIndex > -1 Then
        Dim Button As Byte
        Button = MsgBox("Ban co chac la muôn xoa tu nay?", vbYesNo, "Xoa tu")
        If Button = 6 Then
            Dim i As Long
            For i = lstWord.ListIndex + 1 To n - 1
                a(i - 1) = a(i)
            Next i
            n = n - 1
            ReDim Preserve a(n) As record
            Open "DIC.DAT" For Output As #1
            Lock #1
            Print #1, n
            For i = 0 To n - 1
                With a(i)
                    Print #1, .e
                    Print #1, .e2
                    Print #1, .e3
                    Print #1, .v
                    Print #1, .e4
                    Print #1, .v2
                    Print #1, .v2_
                    Print #1, .v3
                    Print #1, .v3_
                    Print #1, .pre
                    Print #1, .pre2
                End With
            Next i
            Close #1
            With lstWord
                .RemoveItem .ListIndex
            End With
            MsgBox "Da xoa tu", , "Hoan thanh"
            txtWord = ""
            txtExplain = "Giaûi thích"
        End If
    Else
        MsgBox "Ban chua chon tu!", vbExclamation, "Chua chon tu!"
    End If
End Sub

Private Sub cmdFind_Click()
    txtWord = Trim(txtWord)
    Do While InStr(1, txtWord, "  ") > 0
        txtWord = Replace(txtWord, "  ", " ")
    Loop
    Dim i As Long
    For i = 0 To n
        If i = n Then
            MsgBox "Không tìm thây tu """ & txtWord & """!", , "Không tìm thây tu!"
            txtExplain = "Giaûi thích"
            Exit Sub
        End If
        With a(i)
            If LCase(.e) = LCase(txtWord) Then
                txtExplain = .e & vbCrLf & "Töø ñoàng nghóa: "
                If .e2 <> "" Then
                    txtExplain = txtExplain & .e2
                    If .e3 <> "" Then
                        txtExplain = txtExplain & ", " & .e3
                    End If
                Else
                    txtExplain = txtExplain & "khoâng coù"
                End If
                txtExplain = txtExplain & vbCrLf & "Nghóa tieáng Vieät: " & .v & vbCrLf & "Töø traùi nghóa: "
                If .e4 = "" Then
                    txtExplain = txtExplain & "khoâng coù"
                Else
                    txtExplain = txtExplain & .e4
                End If
                txtExplain = txtExplain & vbCrLf & "Quaù khöù (Past tense): "
                If .v2 <> "" Then
                    txtExplain = txtExplain & .v2
                    If .v2_ <> "" Then
                        txtExplain = txtExplain & "/" & .v2_
                    End If
                Else
                    txtExplain = txtExplain & "khoâng coù"
                End If
                txtExplain = txtExplain & vbCrLf & "Ñoäng tính töø quaù khöù (Past participle): "
                If .v3 <> "" Then
                    txtExplain = txtExplain & .v3
                    If .v3_ <> "" Then
                        txtExplain = txtExplain & "/" & .v3_
                    End If
                Else
                    txtExplain = txtExplain & "khoâng coù"
                End If
                txtExplain = txtExplain & vbCrLf & "Giôùi töø ñi keøm: "
                If .pre <> "" Then
                    txtExplain = txtExplain & .pre
                    If .pre2 <> "" Then
                        txtExplain = txtExplain & "/" & .pre2
                    End If
                Else
                    txtExplain = txtExplain & "khoâng coù"
                End If
                Exit Sub
            End If
        End With
    Next i
End Sub

Private Sub Form_Load()
    frmDic.Visible = True
    Open "DIC.DAT" For Input As #1
    Lock #1
    Dim tg As String
    Line Input #1, tg
    n = Val(tg)
    Dim i As Long
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
            lstWord.AddItem .e
        End With
    Next i
    Close #1
End Sub

Private Sub lstWord_Click()
    With lstWord
        txtWord = .List(.ListIndex)
    End With
    cmdFind_Click
End Sub

Private Sub mnuBack_Click()
    Unload frmSet
    Unload frmDic
    Select Case frmMain.frm
        Case 0
            Load frmMain
            frmMain.Visible = True
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
        Case 4
            If frmIrr.lblTime.Caption <> "" Then
                frmIrr.tmrTime.Enabled = True
            End If
        Case 5
            If frmAdj.lblTime.Caption <> "" Then
                frmAdj.tmrTime.Enabled = True
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

Private Sub mnuHelp_Click()
    Unload frmSet
    Unload frmDic
    Load frmHelp
End Sub
Private Sub txtWord_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdFind_Click
    End If
End Sub
