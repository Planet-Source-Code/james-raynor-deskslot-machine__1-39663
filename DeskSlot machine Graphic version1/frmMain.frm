VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actual:"
   ClientHeight    =   1104
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   2664
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":030A
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   3200
      Left            =   1512
      Top             =   1155
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1092
      Top             =   1215
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   636
      Top             =   1245
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   204
      Top             =   1215
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   2820
      Top             =   -60
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1284
      Top             =   1260
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   756
      Top             =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "!"
      Height          =   270
      Left            =   2175
      TabIndex        =   0
      Top             =   420
      Width           =   300
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1272
      Top             =   1320
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   732
      Top             =   1335
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   156
      Top             =   1245
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   144
      Top             =   1185
   End
   Begin VB.Label Hold3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1515
      TabIndex        =   5
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Hold2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   870
      TabIndex        =   4
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Hold1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hold"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   210
      TabIndex        =   3
      Top             =   135
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label txtBonus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No Bonus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   150
      TabIndex        =   2
      Top             =   810
      Width           =   2355
   End
   Begin VB.Label lblVincita 
      BackStyle       =   0  'Transparent
      Caption         =   "+100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2145
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image Img10_Spin 
      Height          =   420
      Left            =   492
      Picture         =   "frmMain.frx":A180
      Top             =   1092
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Image Img10 
      Height          =   420
      Left            =   48
      Picture         =   "frmMain.frx":B22A
      Top             =   1248
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Image Slot3 
      Height          =   420
      Left            =   1488
      Picture         =   "frmMain.frx":C2D4
      Top             =   168
      Width           =   468
   End
   Begin VB.Image Slot2 
      Height          =   420
      Left            =   828
      Picture         =   "frmMain.frx":D37E
      Top             =   168
      Width           =   468
   End
   Begin VB.Image Slot1 
      Height          =   420
      Left            =   168
      Picture         =   "frmMain.frx":E428
      Top             =   168
      Width           =   468
   End
   Begin VB.Image Img9_Spin 
      Height          =   420
      Left            =   7296
      Picture         =   "frmMain.frx":F4D2
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img8_Spin 
      Height          =   420
      Left            =   6768
      Picture         =   "frmMain.frx":1057C
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img7_Spin 
      Height          =   420
      Left            =   6252
      Picture         =   "frmMain.frx":11626
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img6_Spin 
      Height          =   420
      Left            =   5736
      Picture         =   "frmMain.frx":126D0
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img5_Spin 
      Height          =   420
      Left            =   5220
      Picture         =   "frmMain.frx":1377A
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img4_Spin 
      Height          =   420
      Left            =   4716
      Picture         =   "frmMain.frx":14824
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img3_Spin 
      Height          =   420
      Left            =   4188
      Picture         =   "frmMain.frx":158CE
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img2_Spin 
      Height          =   420
      Left            =   3672
      Picture         =   "frmMain.frx":16978
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img1_Spin 
      Height          =   420
      Left            =   3168
      Picture         =   "frmMain.frx":17A22
      Top             =   2520
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Image Img9 
      Height          =   420
      Left            =   7296
      Picture         =   "frmMain.frx":18ACC
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img8 
      Height          =   420
      Left            =   6768
      Picture         =   "frmMain.frx":19B76
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img7 
      Height          =   420
      Left            =   6252
      Picture         =   "frmMain.frx":1AC20
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img6 
      Height          =   420
      Left            =   5736
      Picture         =   "frmMain.frx":1BCCA
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img5 
      Height          =   420
      Left            =   5220
      Picture         =   "frmMain.frx":1CD74
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img4 
      Height          =   420
      Left            =   4716
      Picture         =   "frmMain.frx":1DE1E
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img3 
      Height          =   420
      Left            =   4188
      Picture         =   "frmMain.frx":1EEC8
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img2 
      Height          =   420
      Left            =   3672
      Picture         =   "frmMain.frx":1FF72
      Top             =   2052
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img1 
      Height          =   420
      Left            =   3168
      Picture         =   "frmMain.frx":2101C
      Top             =   2052
      Visible         =   0   'False
      Width           =   468
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim slot1_holded As Boolean: Dim slot2_holded As Boolean: Dim slot3_holded As Boolean
Dim bonustype As Long
Dim scrollCounter As Long
Dim valuetrans As Long
Dim dollari As Long
Dim oldx1 As Long: Dim oldx2 As Long: Dim oldx3 As Long
Dim value1 As Long
Dim value2 As Long
Dim value3 As Long

Private Sub Command1_Click()
dollari = dollari - 1
Me.Caption = "Actual: " & dollari & " $"
Command1.Enabled = False
If bonustype = 10 Then GoTo holdt

Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Exit Sub

holdt:
Hold1.Visible = False: Hold2.Visible = False: Hold3.Visible = False

If slot1_holded = False Then Timer1.Enabled = True: Timer8.Enabled = True
If slot2_holded = False Then Timer2.Enabled = True: Timer9.Enabled = True
If slot3_holded = False Then Timer3.Enabled = True: Timer10.Enabled = True

Timer11.Enabled = True

End Sub

Private Sub Form_DblClick()
If valuetrans = 0 Then MakeTransparent Me.hwnd, 50: valuetrans = 1: Exit Sub
If valuetrans = 1 Then MakeOpaque Me.hwnd: valuetrans = 0: Exit Sub

End Sub

Private Sub Form_Load()
Call PutWindowOnTop(Me)
dollari = 50
Me.Caption = "Actual: " & dollari & " $"
End Sub

Private Sub Slot1_Click()
If bonustype = 10 Then GoTo Hold1
Exit Sub

Hold1:
Command1.Enabled = False


If slot1_holded = False Then slot1_holded = True: Hold1.Visible = True: GoTo usc1
If slot1_holded = True Then slot1_holded = False: Hold1.Visible = False

usc1:
If slot1_holded = True Or slot2_holded = True Or slot3_holded = True Then Command1.Enabled = True

End Sub

Private Sub Slot2_Click()
If bonustype = 10 Then GoTo Hold2
Exit Sub

Hold2:
Command1.Enabled = False

If slot2_holded = False Then slot2_holded = True: Hold2.Visible = True: GoTo usc2
If slot2_holded = True Then slot2_holded = False: Hold2.Visible = False

usc2:
If slot1_holded = True Or slot2_holded = True Or slot3_holded = True Then Command1.Enabled = True


End Sub

Private Sub Slot3_Click()
If bonustype = 10 Then GoTo Hold3
Exit Sub

Hold3:
Command1.Enabled = False

If slot3_holded = False Then slot3_holded = True: Hold3.Visible = True: GoTo usc3
If slot3_holded = True Then slot3_holded = False: Hold3.Visible = False

usc3:
If slot1_holded = True Or slot2_holded = True Or slot3_holded = True Then Command1.Enabled = True


End Sub

Private Sub Timer1_Timer()
Randomize Timer
redox1:
X1 = Int(Rnd * 9) + 1
If oldx1 = X1 Then GoTo redox1
oldx1 = X1

If X1 = 1 Then Slot1.Picture = Img1_Spin.Picture
If X1 = 2 Then Slot1.Picture = Img2_Spin.Picture
If X1 = 3 Then Slot1.Picture = Img3_Spin.Picture
If X1 = 4 Then Slot1.Picture = Img4_Spin.Picture
If X1 = 5 Then Slot1.Picture = Img5_Spin.Picture
If X1 = 6 Then Slot1.Picture = Img6_Spin.Picture
If X1 = 7 Then Slot1.Picture = Img7_Spin.Picture
If X1 = 8 Then Slot1.Picture = Img8_Spin.Picture
If X1 = 9 Then Slot1.Picture = Img9_Spin.Picture

End Sub

Private Sub Timer10_Timer()
Timer10.Enabled = False
Timer3.Enabled = False

Randomize Timer

yy3 = Int(Rnd * 10) + 1
If yy3 = 10 Then GoTo bar3

Y3 = Int(Rnd * 9) + 1

If Y3 = 1 Then Slot3.Picture = Img1.Picture: value3 = 1
If Y3 = 2 Then Slot3.Picture = Img2.Picture: value3 = 2
If Y3 = 3 Then Slot3.Picture = Img3.Picture: value3 = 3
If Y3 = 4 Then Slot3.Picture = Img4.Picture: value3 = 4
If Y3 = 5 Then Slot3.Picture = Img5.Picture: value3 = 5
If Y3 = 6 Then Slot3.Picture = Img6.Picture: value3 = 6
If Y3 = 7 Then Slot3.Picture = Img7.Picture: value3 = 7
If Y3 = 8 Then Slot3.Picture = Img8.Picture: value3 = 8
If Y3 = 9 Then Slot3.Picture = Img9.Picture: value3 = 9
Exit Sub

bar3:
Slot3.Picture = Img10.Picture: value3 = 10


End Sub

Private Sub Timer11_Timer()
Timer11.Enabled = False
Command1.Enabled = True
If value1 = 10 Then dollari = dollari + 1: temporcount = temporcount + 1
If value2 = 10 Then dollari = dollari + 1: temporcount = temporcount + 1
If value3 = 10 Then dollari = dollari + 1: temporcount = temporcount + 1

If value1 = 10 Then If value2 = 10 Then If value3 = 10 Then dollari = dollari + 100: temporcount = temporcount + 100: GoTo refr
If value1 = value2 And value2 = value3 Then dollari = dollari + 30: temporcount = temporcount + 30: GoTo refr
If value1 = value2 Then dollari = dollari + 3: temporcount = temporcount + 3: GoTo refr
If value1 = value3 Then dollari = dollari + 2: temporcount = temporcount + 2: GoTo refr
If value2 = value3 Then dollari = dollari + 2: temporcount = temporcount + 2: GoTo refr

refr:
If bonustype = 1 Then dollari = dollari + temporcount: temporcount = temporcount * 2
If bonustype = 2 Then dollari = dollari + temporcount * 2: temporcount = temporcount * 3
If bonustype = 3 Then dollari = dollari + temporcount * 3: temporcount = temporcount * 4
If bonustype = 4 Then dollari = dollari + temporcount * 4: temporcount = temporcount * 5

Me.Caption = "Actual: " & dollari & " $"
If dollari < 1 Then MsgBox "GAME OVER!", vbCritical: Unload Me
If dollari > 1000 Then MsgBox "CONGRATULATIONS!! YOU BEAT THIS GAME!!", vbExclamation: Unload Me

If temporcount > 0 Then Timer7.Enabled = True: lblVincita.Visible = True: lblVincita.Caption = "+" & temporcount
temporcount = 0

txtBonus.Caption = "No Bonus": bonustype = 0

Randomize Timer: bonus = Int(Rnd * 85) + 1

If bonus = 1 Or bonus = 2 Or bonus = 3 Or bonus = 4 Then bonustype = 1: txtBonus.Caption = "BONUS 2X"
If bonus = 6 Or bonus = 7 Or bonus = 8 Then bonustype = 2: txtBonus.Caption = "BONUS 3X"
If bonus = 9 Or bonus = 10 Then bonustype = 3: txtBonus.Caption = "BONUS 4X"
If bonus = 11 Then bonustype = 4: txtBonus.Caption = "BONUS 5X"
If bonus = 12 Then bonustype = 5: txtBonus.Caption = "BAR BONUS"
If bonus = 13 Or bonus = 14 Or bonus = 15 Then bonustype = 10: txtBonus.Caption = "HOLD TIME!!": Command1.Enabled = False

slot1_holded = False: slot2_holded = False: slot3_holded = False

End Sub

Private Sub Timer2_Timer()

redox2:
X2 = Int(Rnd * 9) + 1
If oldx2 = X2 Then GoTo redox2
oldx2 = X2

If X2 = 1 Then Slot2.Picture = Img1_Spin.Picture
If X2 = 2 Then Slot2.Picture = Img2_Spin.Picture
If X2 = 3 Then Slot2.Picture = Img3_Spin.Picture
If X2 = 4 Then Slot2.Picture = Img4_Spin.Picture
If X2 = 5 Then Slot2.Picture = Img5_Spin.Picture
If X2 = 6 Then Slot2.Picture = Img6_Spin.Picture
If X2 = 7 Then Slot2.Picture = Img7_Spin.Picture
If X2 = 8 Then Slot2.Picture = Img8_Spin.Picture
If X2 = 9 Then Slot2.Picture = Img9_Spin.Picture

End Sub

Private Sub Timer3_Timer()

redox3:
X3 = Int(Rnd * 9) + 1
If oldx3 = X3 Then GoTo redox3
oldx3 = X3

If X3 = 1 Then Slot3.Picture = Img1_Spin.Picture
If X3 = 2 Then Slot3.Picture = Img2_Spin.Picture
If X3 = 3 Then Slot3.Picture = Img3_Spin.Picture
If X3 = 4 Then Slot3.Picture = Img4_Spin.Picture
If X3 = 5 Then Slot3.Picture = Img5_Spin.Picture
If X3 = 6 Then Slot3.Picture = Img6_Spin.Picture
If X3 = 7 Then Slot3.Picture = Img7_Spin.Picture
If X3 = 8 Then Slot3.Picture = Img8_Spin.Picture
If X3 = 9 Then Slot3.Picture = Img9_Spin.Picture

End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Timer1.Enabled = False

Randomize Timer
If bonustype = 5 Then
    yy1 = Int(Rnd * 5) + 1
    If yy1 = 1 Then GoTo bar1
End If

yy1 = Int(Rnd * 10) + 1
If yy1 = 10 Then GoTo bar1

Y1 = Int(Rnd * 9) + 1

If Y1 = 1 Then Slot1.Picture = Img1.Picture: value1 = 1
If Y1 = 2 Then Slot1.Picture = Img2.Picture: value1 = 2
If Y1 = 3 Then Slot1.Picture = Img3.Picture: value1 = 3
If Y1 = 4 Then Slot1.Picture = Img4.Picture: value1 = 4
If Y1 = 5 Then Slot1.Picture = Img5.Picture: value1 = 5
If Y1 = 6 Then Slot1.Picture = Img6.Picture: value1 = 6
If Y1 = 7 Then Slot1.Picture = Img7.Picture: value1 = 7
If Y1 = 8 Then Slot1.Picture = Img8.Picture: value1 = 8
If Y1 = 9 Then Slot1.Picture = Img9.Picture: value1 = 9
Exit Sub

bar1:
Slot1.Picture = Img10.Picture: value1 = 10

End Sub

Private Sub Timer5_Timer()
Timer5.Enabled = False
Timer2.Enabled = False

Randomize Timer
If bonustype = 5 Then
    yy2 = Int(Rnd * 5) + 1
    If yy2 = 1 Then GoTo bar2
End If

yy2 = Int(Rnd * 10) + 1
If yy2 = 10 Then GoTo bar2

Randomize Timer
Y2 = Int(Rnd * 9) + 1

If Y2 = 1 Then Slot2.Picture = Img1.Picture: value2 = 1
If Y2 = 2 Then Slot2.Picture = Img2.Picture: value2 = 2
If Y2 = 3 Then Slot2.Picture = Img3.Picture: value2 = 3
If Y2 = 4 Then Slot2.Picture = Img4.Picture: value2 = 4
If Y2 = 5 Then Slot2.Picture = Img5.Picture: value2 = 5
If Y2 = 6 Then Slot2.Picture = Img6.Picture: value2 = 6
If Y2 = 7 Then Slot2.Picture = Img7.Picture: value2 = 7
If Y2 = 8 Then Slot2.Picture = Img8.Picture: value2 = 8
If Y2 = 9 Then Slot2.Picture = Img9.Picture: value2 = 9
Exit Sub

bar2:
Slot2.Picture = Img10.Picture: value2 = 10


End Sub

Private Sub Timer6_Timer()
Timer6.Enabled = False
Timer3.Enabled = False

Randomize Timer
If bonustype = 5 Then
    yy3 = Int(Rnd * 5) + 1
    If yy3 = 1 Then GoTo bar3
End If

yy3 = Int(Rnd * 10) + 1
If yy3 = 10 Then GoTo bar3

Y3 = Int(Rnd * 9) + 1

If Y3 = 1 Then Slot3.Picture = Img1.Picture: value3 = 1
If Y3 = 2 Then Slot3.Picture = Img2.Picture: value3 = 2
If Y3 = 3 Then Slot3.Picture = Img3.Picture: value3 = 3
If Y3 = 4 Then Slot3.Picture = Img4.Picture: value3 = 4
If Y3 = 5 Then Slot3.Picture = Img5.Picture: value3 = 5
If Y3 = 6 Then Slot3.Picture = Img6.Picture: value3 = 6
If Y3 = 7 Then Slot3.Picture = Img7.Picture: value3 = 7
If Y3 = 8 Then Slot3.Picture = Img8.Picture: value3 = 8
If Y3 = 9 Then Slot3.Picture = Img9.Picture: value3 = 9

bar3:
If yy3 = 10 Then Slot3.Picture = Img10.Picture: value3 = 10

Command1.Enabled = True
If value1 = 10 Then dollari = dollari + 1: temporcount = temporcount + 1
If value2 = 10 Then dollari = dollari + 1: temporcount = temporcount + 1
If value3 = 10 Then dollari = dollari + 1: temporcount = temporcount + 1

If value1 = 10 Then If value2 = 10 Then If value3 = 10 Then dollari = dollari + 100: temporcount = temporcount + 100: GoTo refr
If value1 = value2 And value2 = value3 Then dollari = dollari + 30: temporcount = temporcount + 30: GoTo refr
If value1 = value2 Then dollari = dollari + 3: temporcount = temporcount + 3: GoTo refr
If value1 = value3 Then dollari = dollari + 2: temporcount = temporcount + 2: GoTo refr
If value2 = value3 Then dollari = dollari + 2: temporcount = temporcount + 2: GoTo refr

refr:
If bonustype = 1 Then dollari = dollari + temporcount: temporcount = temporcount * 2
If bonustype = 2 Then dollari = dollari + temporcount * 2: temporcount = temporcount * 3
If bonustype = 3 Then dollari = dollari + temporcount * 3: temporcount = temporcount * 4
If bonustype = 4 Then dollari = dollari + temporcount * 4: temporcount = temporcount * 5

Me.Caption = "Actual: " & dollari & " $"
If dollari < 1 Then MsgBox "GAME OVER!", vbCritical: Unload Me
If dollari > 1000 Then MsgBox "CONGRATULATIONS!! YOU BEAT THIS GAME!!", vbExclamation: Unload Me

If temporcount > 0 Then Timer7.Enabled = True: lblVincita.Visible = True: lblVincita.Caption = "+" & temporcount
temporcount = 0

txtBonus.Caption = "No Bonus": bonustype = 0

Randomize Timer: bonus = Int(Rnd * 85) + 1

If bonus = 1 Or bonus = 2 Or bonus = 3 Or bonus = 4 Then bonustype = 1: txtBonus.Caption = "BONUS 2X"
If bonus = 6 Or bonus = 7 Then bonustype = 2: txtBonus.Caption = "BONUS 3X"
If bonus = 10 Then bonustype = 3: txtBonus.Caption = "BONUS 4X"
If bonus = 11 Then bonustype = 4: txtBonus.Caption = "BONUS 5X"
If bonus = 12 Then bonustype = 5: txtBonus.Caption = "BAR BONUS"
If bonus = 13 Or bonus = 14 Or bonus = 15 Or bonus = 8 Or bonus = 50 Then bonustype = 10: txtBonus.Caption = "HOLD TIME!!": Command1.Enabled = False


End Sub

Private Sub Timer7_Timer()
Timer7.Enabled = False
lblVincita.Visible = False

End Sub

Private Sub Timer8_Timer()
Timer8.Enabled = False
Timer1.Enabled = False

Randomize Timer

yy1 = Int(Rnd * 10) + 1
If yy1 = 10 Then GoTo bar1

Y1 = Int(Rnd * 9) + 1

If Y1 = 1 Then Slot1.Picture = Img1.Picture: value1 = 1
If Y1 = 2 Then Slot1.Picture = Img2.Picture: value1 = 2
If Y1 = 3 Then Slot1.Picture = Img3.Picture: value1 = 3
If Y1 = 4 Then Slot1.Picture = Img4.Picture: value1 = 4
If Y1 = 5 Then Slot1.Picture = Img5.Picture: value1 = 5
If Y1 = 6 Then Slot1.Picture = Img6.Picture: value1 = 6
If Y1 = 7 Then Slot1.Picture = Img7.Picture: value1 = 7
If Y1 = 8 Then Slot1.Picture = Img8.Picture: value1 = 8
If Y1 = 9 Then Slot1.Picture = Img9.Picture: value1 = 9
Exit Sub

bar1:
Slot1.Picture = Img10.Picture: value1 = 10

End Sub

Private Sub Timer9_Timer()
Timer9.Enabled = False
Timer2.Enabled = False

Randomize Timer

yy2 = Int(Rnd * 10) + 1
If yy2 = 10 Then GoTo bar2

Y2 = Int(Rnd * 9) + 1

If Y2 = 1 Then Slot2.Picture = Img1.Picture: value2 = 1
If Y2 = 2 Then Slot2.Picture = Img2.Picture: value2 = 2
If Y2 = 3 Then Slot2.Picture = Img3.Picture: value2 = 3
If Y2 = 4 Then Slot2.Picture = Img4.Picture: value2 = 4
If Y2 = 5 Then Slot2.Picture = Img5.Picture: value2 = 5
If Y2 = 6 Then Slot2.Picture = Img6.Picture: value2 = 6
If Y2 = 7 Then Slot2.Picture = Img7.Picture: value2 = 7
If Y2 = 8 Then Slot2.Picture = Img8.Picture: value2 = 8
If Y2 = 9 Then Slot2.Picture = Img9.Picture: value2 = 9
Exit Sub

bar2:
Slot2.Picture = Img10.Picture: value2 = 10


End Sub
