VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banner Maker Beta 1"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      Caption         =   "Font Underline"
      Height          =   345
      Left            =   4755
      TabIndex        =   25
      Top             =   2025
      Width           =   1140
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Font Italic"
      Height          =   195
      Left            =   4755
      TabIndex        =   24
      Top             =   1755
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Font Bold"
      Height          =   195
      Left            =   4755
      TabIndex        =   23
      Top             =   1485
      Width           =   1005
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1035
      TabIndex        =   22
      Text            =   "Banner Maker"
      Top             =   2385
      Width           =   3540
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4095
      TabIndex        =   20
      Text            =   "100"
      Top             =   3045
      Width           =   405
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2490
      TabIndex        =   19
      Text            =   "5"
      Top             =   3045
      Width           =   405
   End
   Begin VB.PictureBox Picture4 
      Height          =   315
      Left            =   4080
      ScaleHeight     =   255
      ScaleWidth      =   405
      TabIndex        =   16
      Top             =   1500
      Width           =   465
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3480
      TabIndex        =   15
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1020
      TabIndex        =   14
      Text            =   "C:\MyBanner.bmp"
      Top             =   3525
      Width           =   2400
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&About"
      Height          =   350
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4020
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      Height          =   350
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4020
      Width           =   870
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "...."
      Height          =   350
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3045
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   4815
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   9
      Top             =   3210
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   1920
      Width           =   1620
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1005
      Picture         =   "Form1.frx":3042
      ScaleHeight     =   255
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   1500
      Width           =   3060
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   915
      Left            =   75
      ScaleHeight     =   855
      ScaleWidth      =   5790
      TabIndex        =   2
      Top             =   150
      Width           =   5850
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save"
      Height          =   350
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4020
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Make"
      Height          =   350
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4020
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Banner Title"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   2445
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Title Alignment"
      Height          =   195
      Index           =   6
      Left            =   3000
      TabIndex        =   18
      Top             =   3105
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Title Top"
      Height          =   195
      Index           =   5
      Left            =   1785
      TabIndex        =   17
      Top             =   3105
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Save image"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   3570
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BackGround"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3105
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   30
      X2              =   1260
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   30
      X2              =   1260
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Font Size"
      Height          =   195
      Index           =   2
      Left            =   2745
      TabIndex        =   7
      Top             =   1995
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Font Name"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1995
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Font Colour"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1545
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Sub CenterForm(Frm As Form)
With Frm
    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2
End With

End Sub
Sub MakeBanner(Xpos, YPos As Integer, Title, TitleFont As String, TextSize As Integer, bsBold, bsItalic, bsUnLine, TextColour As Long)
Dim X, Y, I, J As Single

    I = Picture1.ScaleWidth
    J = Picture1.ScaleHeight
    
    Y = 0
    Do While Y < Picture2.ScaleHeight
        X = 0
        Do While X < Picture2.ScaleWidth
            Picture2.PaintPicture Picture1.Picture, X, Y, I, J
            Picture2.FontName = TitleFont
            Picture2.FontBold = bsBold
            Picture2.FontItalic = bsItalic
            Picture2.FontUnderline = bsUnLine
            Picture2.ForeColor = TextColour
            Picture2.FontSize = TextSize
            TextOut Picture2.hdc, Xpos, YPos, Title, Len(Title)
            X = X + I
        Loop
        Y = Y + J
    Loop
    
End Sub


Private Sub Command1_Click()
MakeBanner Val(Text3), Val(Text2), Text4, Combo1.Text, Val(Combo2), _
Check1.Value, Check2.Value, Check3.Value, Picture4.BackColor

End Sub

Private Sub Command2_Click()
SavePicture Picture2.Image, Text1.Text
 MsgBox "Your Banner have been saved to " & Text1.Text
 
End Sub

Private Sub Command3_Click()
Form2.Show
 Form1.Hide
 
End Sub





Private Sub Command4_Click()
MsgBox "More of my projects at http://www.codearchive.com/~dreamvb/": End

End Sub

Private Sub Command5_Click()
 MsgBox "Banner Maker by Ben Jones Please Vote if you like it", vbInformation
 
End Sub

Private Sub Form_Load()
CenterForm Form1

  For I = 1 To Screen.FontCount - 1
   Combo1.AddItem Screen.Fonts(I)
    Next
 
 For J = 10 To 40 Step 2
    Combo2.AddItem J
 Next
 Combo1.ListIndex = 4
 Combo2.ListIndex = 8
 Picture4.BackColor = vbRed
 
End Sub

Private Sub Form_Resize()
 Line1(0).X2 = Me.Width
 Line1(1).X2 = Me.Width
 
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim MSG
On Error Resume Next
 
 If Button = 1 Then
    Picture4.BackColor = Picture3.Point(X, Y)
 If Err Then Err.Clear
 
 End If
 
End Sub

Private Sub Picture5_Click()
Picture1.Picture = Picture5.Picture

End Sub

