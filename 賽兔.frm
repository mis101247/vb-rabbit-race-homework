VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "2009年國際世界大賽-肥兔100M"
   ClientHeight    =   6870
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   Picture         =   "賽兔.frx":0000
   ScaleHeight     =   6870
   ScaleWidth      =   10080
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton Command3 
      Caption         =   "連勝說明"
      Height          =   481
      Left            =   3393
      MaskColor       =   &H8000000F&
      TabIndex        =   25
      Top             =   6318
      Width           =   598
   End
   Begin VB.TextBox money1 
      Enabled         =   0   'False
      Height          =   364
      Left            =   6675
      TabIndex        =   23
      Text            =   "50000"
      Top             =   5967
      Width           =   1665
   End
   Begin VB.TextBox money2 
      Height          =   364
      Left            =   6675
      TabIndex        =   21
      Text            =   "0"
      ToolTipText     =   "贏的話賺6倍錢"
      Top             =   6318
      Width           =   1665
   End
   Begin VB.OptionButton Op 
      BackColor       =   &H00C0FFFF&
      Height          =   286
      Index           =   4
      Left            =   234
      TabIndex        =   6
      Top             =   4914
      UseMaskColor    =   -1  'True
      Width           =   247
   End
   Begin VB.OptionButton Op 
      BackColor       =   &H00C0FFFF&
      Height          =   286
      Index           =   3
      Left            =   234
      TabIndex        =   5
      Top             =   3861
      UseMaskColor    =   -1  'True
      Width           =   247
   End
   Begin VB.OptionButton Op 
      BackColor       =   &H00C0FFFF&
      Height          =   286
      Index           =   2
      Left            =   234
      TabIndex        =   4
      Top             =   2808
      UseMaskColor    =   -1  'True
      Width           =   247
   End
   Begin VB.OptionButton Op 
      BackColor       =   &H00C0FFFF&
      Height          =   286
      Index           =   1
      Left            =   234
      TabIndex        =   3
      Top             =   1755
      UseMaskColor    =   -1  'True
      Width           =   247
   End
   Begin VB.CommandButton Command2 
      Caption         =   "預備"
      Height          =   481
      Left            =   234
      TabIndex        =   2
      Top             =   6090
      Width           =   1066
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   481
      Left            =   234
      TabIndex        =   1
      Top             =   5625
      Width           =   1066
   End
   Begin VB.OptionButton Op 
      BackColor       =   &H00C0FFFF&
      Height          =   286
      Index           =   0
      Left            =   234
      TabIndex        =   0
      Top             =   702
      UseMaskColor    =   -1  'True
      Width           =   247
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   8658
      Top             =   6318
   End
   Begin VB.Label CwinL 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Left            =   2691
      TabIndex        =   24
      Top             =   6318
      Width           =   1066
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "籌碼餘額(0為輸):"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4335
      TabIndex        =   22
      Top             =   5970
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "下注金額(賠率6):"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4335
      TabIndex        =   20
      Top             =   6315
      Width           =   2595
   End
   Begin VB.Label WL 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Index           =   1
      Left            =   2691
      TabIndex        =   19
      Top             =   5967
      Width           =   1066
   End
   Begin VB.Label WL 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   481
      Index           =   0
      Left            =   1521
      TabIndex        =   18
      Top             =   5967
      Width           =   1066
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "玩家      贏:   輸:  連勝次數:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1290
      Left            =   1440
      TabIndex        =   17
      Top             =   5610
      Width           =   1995
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      Index           =   5
      X1              =   117
      X2              =   9360
      Y1              =   5499
      Y2              =   5499
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      Index           =   4
      X1              =   120
      X2              =   9363
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      Index           =   3
      X1              =   120
      X2              =   9363
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      Index           =   2
      X1              =   120
      X2              =   9363
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      Index           =   1
      X1              =   120
      X2              =   9363
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H008080FF&
      Index           =   0
      X1              =   117
      X2              =   9360
      Y1              =   117
      Y2              =   117
   End
   Begin VB.Image R2 
      Height          =   1065
      Index           =   1
      Left            =   9015
      Picture         =   "賽兔.frx":9080
      Top             =   5730
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image R2 
      Height          =   1050
      Index           =   0
      Left            =   9015
      Picture         =   "賽兔.frx":93EA
      Top             =   5730
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   4
      Left            =   480
      Picture         =   "賽兔.frx":9747
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   3
      Left            =   480
      Picture         =   "賽兔.frx":9AD5
      Top             =   3390
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   2
      Left            =   480
      Picture         =   "賽兔.frx":9E63
      Top             =   2295
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   1
      Left            =   480
      Picture         =   "賽兔.frx":A1F1
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   0
      Left            =   480
      Picture         =   "賽兔.frx":A57F
      Top             =   120
      Width           =   765
   End
   Begin VB.Line endL 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   9360
      X2              =   9360
      Y1              =   117
      Y2              =   5517
   End
   Begin VB.Label Score 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   7485
      TabIndex        =   8
      Top             =   450
      Width           =   1065
   End
   Begin VB.Label No 
      BackStyle       =   0  '透明
      Caption         =   "1號選手-肥龍兔-贏的次數:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Index           =   0
      Left            =   1635
      TabIndex        =   7
      Top             =   450
      Width           =   5985
   End
   Begin VB.Label Score 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   7485
      TabIndex        =   10
      Top             =   1530
      Width           =   1065
   End
   Begin VB.Label No 
      BackStyle       =   0  '透明
      Caption         =   "2號選手-霸王兔-贏的次數:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Index           =   1
      Left            =   1635
      TabIndex        =   9
      Top             =   1530
      Width           =   5985
   End
   Begin VB.Label Score 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   7485
      TabIndex        =   12
      Top             =   2640
      Width           =   1065
   End
   Begin VB.Label No 
      BackStyle       =   0  '透明
      Caption         =   "3號選手-OPEN兔-贏的次數:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Index           =   2
      Left            =   1635
      TabIndex        =   11
      Top             =   2640
      Width           =   6105
   End
   Begin VB.Label Score 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   7440
      TabIndex        =   14
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Label No 
      BackStyle       =   0  '透明
      Caption         =   "4號選手-耍劍兔-贏的次數:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Index           =   3
      Left            =   1635
      TabIndex        =   13
      Top             =   3630
      Width           =   5985
   End
   Begin VB.Label Score 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   25.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   7485
      TabIndex        =   16
      Top             =   4800
      Width           =   1065
   End
   Begin VB.Label No 
      BackStyle       =   0  '透明
      Caption         =   "5號選手-食人兔-贏的次數:"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Index           =   4
      Left            =   1635
      TabIndex        =   15
      Top             =   4800
      Width           =   5985
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   0
      Left            =   480
      Picture         =   "賽兔.frx":A90D
      Top             =   120
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   1
      Left            =   480
      Picture         =   "賽兔.frx":AB5A
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   2
      Left            =   480
      Picture         =   "賽兔.frx":ADA7
      Top             =   2295
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   3
      Left            =   480
      Picture         =   "賽兔.frx":AFF4
      Top             =   3390
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   4
      Left            =   480
      Picture         =   "賽兔.frx":B241
      Top             =   4455
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  '不透明
      BorderColor     =   &H008080FF&
      Height          =   5400
      Left            =   120
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nextp As Integer '定義圖片下一張位子
Dim Cwin As Integer '連勝計次
Dim money As Double '籌碼金額

Private Sub command1_click()
If money2 <= money And money2 >= 0 Then '下注金額為正數並小餘籌碼金額亦執行
  For i = 0 To 4
  Op(i).Enabled = False                                      '關閉OptionButton選取
  Next i
Timer1.Enabled = True                                    '開啟Timer
Command1.Enabled = False                            '關閉 開始按鍵
Command2.Enabled = False                            '關閉 預備按鍵
money2.Locked = True                                    '關閉  下注金額內容修改
Else
MsgBox "下注金額錯誤or你已經沒有錢了", , "錯誤"
End If
End Sub

Private Sub Command2_Click()
For i = 0 To 4
R1(i).Left = Image1(i).Left                                             '賤兔圖片拉回原點
Op(i).Enabled = True                                      '開啟OptionButton選取
Next i
Command1.Enabled = True
Command2.Enabled = True
money2.Locked = False
End Sub

Private Sub Command3_Click()
MsgBox "目前已贏 " & Cwin & " 次-連勝破0次-押的選手士氣+5%", , "說明"
End Sub

Private Sub Form_Load()
money = 50000 '設定原有五萬
End Sub


Private Sub Timer1_Timer()

'Randomize
    nextp = (nextp + 1) Mod 2                      '遞減 先1後0
     For i = 0 To 4
     
    R1(i).Picture = R2(nextp).Picture
    
    If Cwin < 2 Then                                   '連勝小於1~小於連勝計次2
     X = Int(Rnd * 500) + 1
     R1(i).Left = R1(i).Left + X
     Else
     X = Int(Rnd * 500) + 1
     R1(i).Left = R1(i).Left + X
      For j = 0 To 4
           If Op(j).Value = True Then
           X = Int(Rnd * 500) + 1
           R1(j).Left = R1(j).Left + (X / 20) ' 亂數+亂數的20分之1
           End If
       Next j
     End If
     
        If R1(i).Left + R1(i).Width >= endL.X1 Then  '圖片左線位子+圖片寬度>=終點線
         Timer1.Enabled = False
         Command2.Enabled = True
         Score(i).Caption = Score(i).Caption + 1
             If Op(i).Value = True Then
               WL(0) = WL(0) + 1
               Cwin = Cwin + 1
               money = money1 + money2 * 6
               money1 = money
               MsgBox ("你贏了"), , "You win!!"
                If Cwin >= 2 Then CwinL.Caption = (Cwin - 1)
             Else
               WL(1) = WL(1) + 1
               Cwin = 0
               CwinL.Caption = 0
               money = money1 - money2
               money1 = money
               MsgBox ("你輸了"), , "You lose!!"
              End If
          Exit For
          End If
    Next i

End Sub


