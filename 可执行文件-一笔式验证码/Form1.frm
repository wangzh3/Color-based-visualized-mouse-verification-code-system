VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "注册"
   ClientHeight    =   5868
   ClientLeft      =   5712
   ClientTop       =   2976
   ClientWidth     =   6648
   LinkTopic       =   "Form1"
   ScaleHeight     =   5868
   ScaleWidth      =   6648
   Begin VB.CommandButton Command5 
      Caption         =   "显示密码字符"
      Height          =   372
      Left            =   4080
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.CommandButton Command4 
      Caption         =   "更换图形"
      Height          =   400
      Left            =   4080
      TabIndex        =   19
      Top             =   3480
      Visible         =   0   'False
      Width           =   1330
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   1
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   3
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   4
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   6
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   7
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CheckBox Check 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Index           =   8
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "图形选择完成"
      Height          =   400
      Left            =   4080
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   1330
   End
   Begin VB.CommandButton Command2 
      Caption         =   "注册密码完成"
      Height          =   400
      Left            =   4080
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1330
   End
   Begin VB.TextBox Text3 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   400
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "用户名确认"
      Height          =   400
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   1330
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "形状及颜色完全相同的所有图形"
      Height          =   492
      Left            =   3600
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   3000
      Top             =   2280
      Width           =   492
   End
   Begin VB.Label Label5 
      Caption         =   "请从下列图形中选择所有的与"
      Height          =   372
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.Label Label4 
      Caption         =   "再次输入密码"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "输入注册密码"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "输入注册用户名"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(8), tu, xzt As String
Dim t1(8), t2(8) As Integer
Dim t As Integer
Private Sub Command1_Click()
Open "user_list" For Input As #1
ReDim u(i) As user
flage = 0
i = 0
Do While Not EOF(1)
    Input #1, u(i).user_1, u(i).user_2
    If u(i).user_1 = Text1.Text Then
        x = "对不起，用户名" + Text1.Text + "已注册，请重新输入！"
        xx = MsgBox(x, 64, "注册提示")
        Text1.Text = ""
        Text1.SetFocus
        flage = 1
        
        Exit Do
    End If
Loop
Close #1
If flage = 0 Then
    xx = MsgBox("用户名申请成功，继续注册其他信息！", 0, "注册信息")
    Label3.Visible = True
    Label4.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Command2.Visible = True
    Command5.Visible = True
    Command1.Enabled = False
    Text2.SetFocus
End If

End Sub

Private Sub Command2_Click()
If Text2.Text <> Text3.Text Then
    xx = MsgBox("两次输入密码不统一，请重新输入！", 64, "密码信息")
    Text2.Text = ""
    Text3.Text = ""
    Text2.SetFocus
Else
    Command2.Enabled = False
    Label5.Visible = True
    Label2.Visible = True
    Command3.Visible = True
    Command4.Visible = True
    Randomize (Second(Now))
    
    For i = 0 To 8
        Check(i).Visible = True
        Check(i).Value = 0
        t1(i) = Int(Rnd * 4 + 1)
        t2(i) = Int(Rnd * 5 + 1)
        a(i) = Trim(Str(t1(i))) + Trim(Str(t2(i)))
        tu = a(i) + ".jpg"
        Check(i).Picture = LoadPicture(tu)
    Next
    t = Int(Rnd * 9)
    tu = a(t) + ".jpg"
    xzt = a(t)
    Image1.Picture = LoadPicture(tu)
    

End If
End Sub

Private Sub Command3_Click()
flage = 0
For i = 0 To 8
    
    If Check(i).Value = 1 Then
        If a(i) <> xzt Then
            flage = 1
            
            Exit For
        End If
    Else
        If a(i) = xzt Then
            flage = 1
            Exit For
        End If
    End If
Next
If flage = 0 Then
    xx = MsgBox("      注册完成！     ", 0, "图形信息提示")
    Open "user_list" For Append As #1
    Write #1, Text1.Text, Text2.Text
    Close #1
    End
Else
    xx = MsgBox("识别图形有误，请重新进行图形识别！", 64, "图形信息提示")
    For i = 0 To 8
        Check(i).Value = 0
        t1(i) = Int(Rnd * 4 + 1)
        t2(i) = Int(Rnd * 5 + 1)
        a(i) = Trim(Str(t1(i))) + Trim(Str(t2(i)))
        tu = a(i) + ".jpg"
        Check(i).Picture = LoadPicture(tu)
    Next
    t = Int(Rnd * 9)
    tu = a(t) + ".jpg"
    xzt = a(t)
    Image1.Picture = LoadPicture(tu)
End If

End Sub

Private Sub Command4_Click()
 For i = 0 To 8
        Check(i).Value = 0
        t1(i) = Int(Rnd * 4 + 1)
        t2(i) = Int(Rnd * 5 + 1)
        a(i) = Trim(Str(t1(i))) + Trim(Str(t2(i)))
        tu = a(i) + ".jpg"
        Check(i).Picture = LoadPicture(tu)
    Next
    t = Int(Rnd * 9)
    tu = a(t) + ".jpg"
    xzt = a(t)

    Image1.Picture = LoadPicture(tu)
End Sub

Private Sub Command5_Click()
If Command5.Caption = "显示密码字符" Then
    Text2.PasswordChar = ""
    Text3.PasswordChar = ""
    Command5.Caption = "隐藏密码字符"
Else
    Text2.PasswordChar = "*"
    Text3.PasswordChar = "*"
    Command5.Caption = "显示密码字符"
End If

End Sub

