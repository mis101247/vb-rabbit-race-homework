VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "2009�~��ڥ@�ɤj��-�Ψ�100M"
   ClientHeight    =   6870
   ClientLeft      =   105
   ClientTop       =   390
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   Picture         =   "�ɨ�.frx":0000
   ScaleHeight     =   6870
   ScaleWidth      =   10080
   StartUpPosition =   2  '�ù�����
   Begin VB.CommandButton Command3 
      Caption         =   "�s�ӻ���"
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
      ToolTipText     =   "Ĺ������6����"
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
      Caption         =   "�w��"
      Height          =   481
      Left            =   234
      TabIndex        =   2
      Top             =   6090
      Width           =   1066
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�}�l"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "�w�X�l�B(0����):"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "�U�`���B(�߲v6):"
      BeginProperty Font 
         Name            =   "�з���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "���a      Ĺ:   ��:  �s�Ӧ���:"
      BeginProperty Font 
         Name            =   "�з���"
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
      Picture         =   "�ɨ�.frx":9080
      Top             =   5730
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image R2 
      Height          =   1050
      Index           =   0
      Left            =   9015
      Picture         =   "�ɨ�.frx":93EA
      Top             =   5730
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   4
      Left            =   480
      Picture         =   "�ɨ�.frx":9747
      Top             =   4440
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   3
      Left            =   480
      Picture         =   "�ɨ�.frx":9AD5
      Top             =   3390
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   2
      Left            =   480
      Picture         =   "�ɨ�.frx":9E63
      Top             =   2295
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   1
      Left            =   480
      Picture         =   "�ɨ�.frx":A1F1
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image R1 
      Height          =   1065
      Index           =   0
      Left            =   480
      Picture         =   "�ɨ�.frx":A57F
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "1�����-���s��-Ĺ������:"
      BeginProperty Font 
         Name            =   "�з���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "2�����-�Q����-Ĺ������:"
      BeginProperty Font 
         Name            =   "�з���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "3�����-OPEN��-Ĺ������:"
      BeginProperty Font 
         Name            =   "�з���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "4�����-�A�C��-Ĺ������:"
      BeginProperty Font 
         Name            =   "�з���"
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
      Alignment       =   2  '�m�����
      BackStyle       =   0  '�z��
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�з���"
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
      BackStyle       =   0  '�z��
      Caption         =   "5�����-���H��-Ĺ������:"
      BeginProperty Font 
         Name            =   "�з���"
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
      Picture         =   "�ɨ�.frx":A90D
      Top             =   120
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   1
      Left            =   480
      Picture         =   "�ɨ�.frx":AB5A
      Top             =   1200
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   2
      Left            =   480
      Picture         =   "�ɨ�.frx":ADA7
      Top             =   2295
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   3
      Left            =   480
      Picture         =   "�ɨ�.frx":AFF4
      Top             =   3390
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   1065
      Index           =   4
      Left            =   480
      Picture         =   "�ɨ�.frx":B241
      Top             =   4455
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  '���z��
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
Dim nextp As Integer '�w�q�Ϥ��U�@�i��l
Dim Cwin As Integer '�s�ӭp��
Dim money As Double '�w�X���B

Private Sub command1_click()
If money2 <= money And money2 >= 0 Then '�U�`���B�����ƨäp�l�w�X���B�����
  For i = 0 To 4
  Op(i).Enabled = False                                      '����OptionButton���
  Next i
Timer1.Enabled = True                                    '�}��Timer
Command1.Enabled = False                            '���� �}�l����
Command2.Enabled = False                            '���� �w�ƫ���
money2.Locked = True                                    '����  �U�`���B���e�ק�
Else
MsgBox "�U�`���B���~or�A�w�g�S�����F", , "���~"
End If
End Sub

Private Sub Command2_Click()
For i = 0 To 4
R1(i).Left = Image1(i).Left                                             '��߹Ϥ��Ԧ^���I
Op(i).Enabled = True                                      '�}��OptionButton���
Next i
Command1.Enabled = True
Command2.Enabled = True
money2.Locked = False
End Sub

Private Sub Command3_Click()
MsgBox "�ثe�wĹ " & Cwin & " ��-�s�ӯ}0��-�㪺���h��+5%", , "����"
End Sub

Private Sub Form_Load()
money = 50000 '�]�w�즳���U
End Sub


Private Sub Timer1_Timer()

'Randomize
    nextp = (nextp + 1) Mod 2                      '���� ��1��0
     For i = 0 To 4
     
    R1(i).Picture = R2(nextp).Picture
    
    If Cwin < 2 Then                                   '�s�Ӥp��1~�p��s�ӭp��2
     X = Int(Rnd * 500) + 1
     R1(i).Left = R1(i).Left + X
     Else
     X = Int(Rnd * 500) + 1
     R1(i).Left = R1(i).Left + X
      For j = 0 To 4
           If Op(j).Value = True Then
           X = Int(Rnd * 500) + 1
           R1(j).Left = R1(j).Left + (X / 20) ' �ü�+�üƪ�20����1
           End If
       Next j
     End If
     
        If R1(i).Left + R1(i).Width >= endL.X1 Then  '�Ϥ����u��l+�Ϥ��e��>=���I�u
         Timer1.Enabled = False
         Command2.Enabled = True
         Score(i).Caption = Score(i).Caption + 1
             If Op(i).Value = True Then
               WL(0) = WL(0) + 1
               Cwin = Cwin + 1
               money = money1 + money2 * 6
               money1 = money
               MsgBox ("�AĹ�F"), , "You win!!"
                If Cwin >= 2 Then CwinL.Caption = (Cwin - 1)
             Else
               WL(1) = WL(1) + 1
               Cwin = 0
               CwinL.Caption = 0
               money = money1 - money2
               money1 = money
               MsgBox ("�A��F"), , "You lose!!"
              End If
          Exit For
          End If
    Next i

End Sub


