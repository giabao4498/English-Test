VERSION 5.00
Begin VB.Form frmSet 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   825
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "VNI-Times"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Huûy"
      Height          =   495
      Left            =   6480
      TabIndex        =   24
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton cmdDone 
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtPre2 
      Height          =   540
      Left            =   3480
      TabIndex        =   22
      Top             =   8160
      Width           =   2895
   End
   Begin VB.TextBox txtPre 
      Height          =   540
      Left            =   3480
      TabIndex        =   21
      Top             =   7440
      Width           =   2895
   End
   Begin VB.TextBox txtV3_ 
      Height          =   540
      Left            =   3480
      TabIndex        =   20
      Top             =   6765
      Width           =   2895
   End
   Begin VB.TextBox txtV3 
      Height          =   540
      Left            =   3480
      TabIndex        =   19
      Top             =   6045
      Width           =   2895
   End
   Begin VB.TextBox txtV2_ 
      Height          =   540
      Left            =   3480
      TabIndex        =   18
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox txtV2 
      Height          =   540
      Left            =   3480
      TabIndex        =   17
      Top             =   4605
      Width           =   2895
   End
   Begin VB.TextBox txtE4 
      Height          =   540
      Left            =   3480
      TabIndex        =   16
      Top             =   3885
      Width           =   2895
   End
   Begin VB.TextBox txtE3 
      Height          =   540
      Left            =   3480
      TabIndex        =   15
      Top             =   3120
      Width           =   2895
   End
   Begin VB.TextBox txtE2 
      Height          =   540
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtV 
      Height          =   540
      Left            =   3480
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtE 
      Height          =   540
      Left            =   3480
      TabIndex        =   2
      Top             =   915
      Width           =   2895
   End
   Begin VB.Label lblPre2 
      Alignment       =   2  'Center
      Caption         =   "Giôùi töø ñi keøm 2:"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   8160
      Width           =   3255
   End
   Begin VB.Label lblPre 
      Alignment       =   2  'Center
      Caption         =   "Giôùi töø ñi keøm 1:"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label lblV3_ 
      Alignment       =   2  'Center
      Caption         =   "Ñoäng tính töø quaù khöù 2:"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label lblV3 
      Alignment       =   2  'Center
      Caption         =   "Ñoäng tính töø quaù khöù 1:"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   3255
   End
   Begin VB.Label lblV2_ 
      Alignment       =   2  'Center
      Caption         =   "Quaù khöù 2:"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label lblV2 
      Alignment       =   2  'Center
      Caption         =   "Quaù khöù:"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label lblE4 
      Alignment       =   2  'Center
      Caption         =   "Töø traùi nghóa:"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label lblE3 
      Alignment       =   2  'Center
      Caption         =   "Töø ñoàng nghóa 2:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label lblE2 
      Alignment       =   2  'Center
      Caption         =   "Töø ñoàng nghóa 1:"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lblV 
      Alignment       =   2  'Center
      Caption         =   "Töø tieáng Vieät:"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblE 
      Alignment       =   2  'Center
      Caption         =   "Töø tieáng Anh:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public b As String

Private Sub cmdCancel_Click()
    Unload frmSet
End Sub

Private Sub cmdDone_Click()
    Dim kt As Boolean
    kt = True
    If txtE = "" Then
        MsgBox "Ban chua nhâp tu tiêng Anh!", vbExclamation, "Du liêu không hop lê!"
        kt = False
    End If
    If txtV = "" Then
        MsgBox "Ban chua nhâp tu tiêng Viêt!", vbExclamation, "Du liêu không hop lê!"
        kt = False
    End If
    If (txtE2 = "") And (txtE3 <> "") Then
        txtE2 = txtE3
        txtE3 = ""
    End If
    If (txtV2 = "") And (txtV2_ <> "") Then
        txtV2 = txtV2_
        txtV2_ = ""
    End If
    If (txtV3 = "") And (txtV3_ <> "") Then
        txtV3 = txtV3_
        txtV3_ = ""
    End If
    If (txtV2 = "") And (txtV3 <> "") Then
        MsgBox "Ban chua nhâp dông tu qua khu!", vbExclamation, "Du liêu không hop lê!"
        kt = False
    End If
    If (txtV3 = "") And (txtV2 <> "") Then
        MsgBox "Ban chua nhâp dông tính tu qua khu!", vbExclamation, "Du liêu không hop lê!"
        kt = False
    End If
    If (txtPre = "") And (txtPre2 <> "") Then
        txtPre = txtPre2
        txtPre2 = ""
    End If
    If Not kt Then
        Exit Sub
    Else
        If cmdDone.Caption = "Chænh söûa" Then
            frmDic.change
        Else
            frmDic.add
        End If
        Unload frmSet
    End If
End Sub

Private Sub Form_Load()
    frmSet.Visible = True
End Sub
