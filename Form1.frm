VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PC organ!"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer PlayTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4680
      Top             =   3840
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   3960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Min             =   1
      Max             =   1000
      SelStart        =   250
      TickFrequency   =   50
      Value           =   250
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Playback"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pause note"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   10035
      TabIndex        =   1
      Top             =   0
      Width           =   10095
      Begin VB.Line Play 
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Shape PNote 
         BorderColor     =   &H80000007&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2160
         Top             =   -200
         Width           =   135
      End
      Begin VB.Shape Note 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   360
         Shape           =   3  'Circle
         Top             =   -200
         Width           =   135
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   10080
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   10110
      TabIndex        =   0
      Top             =   1800
      Width           =   10170
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   41
         Left            =   7800
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   40
         Left            =   8160
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   38
         Left            =   8880
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   37
         Left            =   9240
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   36
         Left            =   9600
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   1
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   3
         Left            =   600
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   6
         Left            =   1320
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   8
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   10
         Left            =   2040
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   13
         Left            =   2760
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   15
         Left            =   3120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   18
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   20
         Left            =   4200
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   22
         Left            =   4560
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   0
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   2
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   4
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   5
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   7
         Left            =   1440
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   9
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   11
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   12
         Left            =   2520
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   14
         Left            =   2880
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   16
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   17
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   19
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   21
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   23
         Left            =   4680
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   29
         Left            =   5280
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   28
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   26
         Left            =   6360
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   25
         Left            =   6720
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   24
         Left            =   7080
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   30
         Left            =   5040
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   31
         Left            =   5400
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   32
         Left            =   5760
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   27
         Left            =   6120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   33
         Left            =   6480
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   34
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   35
         Left            =   7200
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   47
         Left            =   9720
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   46
         Left            =   9360
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   45
         Left            =   9000
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   39
         Left            =   8640
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   44
         Left            =   8280
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   43
         Left            =   7920
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
      Begin VB.Shape Key 
         BorderColor     =   &H80000006&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1575
         Index           =   42
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo:"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   3600
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
For i = 1 To Note.UBound
Unload Note(i)
Next i
StepStart = 50
Picture2.SetFocus
End Sub

Private Sub Command2_Click()
Unload Note(Note.UBound)
StepStart = StepStart - StepF
If Note.UBound = 0 Then Command2.Enabled = False
Picture2.SetFocus
End Sub

Private Sub Command3_Click()
Load Note(Note.UBound + 1)
StepStart = StepStart + StepF

Note(Note.UBound).Shape = 0
Note(Note.UBound).FillColor = &HFFFF&
Note(Note.UBound).Top = (Picture2.Height / DivN) - Note(Note.UBound).Height / 2
Note(Note.UBound).Left = StepStart
Note(Note.UBound).Visible = True
Picture2.SetFocus
End Sub

Private Sub Command4_Click()
Play.X1 = 50 + StepF
Play.X2 = 50 + StepF
PlayTimer.Enabled = True
End Sub

Private Sub Form_Activate()
Picture2.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
MsgBox KeyCode
End Sub

Private Sub Form_Load()
Tempo = 250
StepStart = 50
Picture2.Height = Picture1.Height
For i = 1 To 13
Load Line1(i)
Line1(i).Y1 = Int(Picture2.Height / DivN) * i
Line1(i).Y2 = Int(Picture2.Height / DivN) * i
Line1(i).Visible = True
Next i
'Picture2.SetFocus
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.Enabled = True

If Y > (Picture1.Height / 2) + 170 Then
Select Case X

Case 10 To Picture1.Width / 28
Beep C1, Dur
PlaceNote 0&, 13

Case Picture1.Width / 28 To Picture1.Width / 28 * 2
Beep D1, Dur
PlaceNote 0&, 11

Case Picture1.Width / 28 * 2 To Picture1.Width / 28 * 3
Beep E1, Dur
PlaceNote 0&, 9

Case Picture1.Width / 28 * 3 To Picture1.Width / 28 * 4
Beep F1, Dur
PlaceNote 0&, 8

Case Picture1.Width / 28 * 4 To Picture1.Width / 28 * 5
Beep G1, Dur
PlaceNote 0&, 6

Case Picture1.Width / 28 * 5 To Picture1.Width / 28 * 6
Beep A1, Dur
PlaceNote 0&, 4

Case Picture1.Width / 28 * 6 To Picture1.Width / 28 * 7
Beep B1, Dur
PlaceNote 0&, 2

Case Picture1.Width / 28 * 7 To Picture1.Width / 28 * 8
Beep (C1 * 2), Dur
PlaceNote Blue, 13

Case Picture1.Width / 28 * 8 To Picture1.Width / 28 * 9
Beep (D1 * 2), Dur
PlaceNote Blue, 11

Case Picture1.Width / 28 * 9 To Picture1.Width / 28 * 10
Beep (E1 * 2), Dur
PlaceNote Blue, 9

Case Picture1.Width / 28 * 10 To Picture1.Width / 28 * 11
Beep (F1 * 2), Dur
PlaceNote Blue, 8

Case Picture1.Width / 28 * 11 To Picture1.Width / 28 * 12
Beep (G1 * 2), Dur
PlaceNote Blue, 6

Case Picture1.Width / 28 * 12 To Picture1.Width / 28 * 13
Beep (A1 * 2), Dur
PlaceNote Blue, 4

Case Picture1.Width / 28 * 13 To Picture1.Width / 28 * 14
Beep (B1 * 2), Dur
PlaceNote Blue, 2

Case Picture1.Width / 28 * 14 To Picture1.Width / 28 * 15
Beep (C1 * 4), Dur
PlaceNote Green, 13

Case Picture1.Width / 28 * 15 To Picture1.Width / 28 * 16
Beep (D1 * 4), Dur
PlaceNote Green, 11

Case Picture1.Width / 28 * 16 To Picture1.Width / 28 * 17
Beep (E1 * 4), Dur
PlaceNote Green, 9

Case Picture1.Width / 28 * 17 To Picture1.Width / 28 * 18
Beep (F1 * 4), Dur
PlaceNote Green, 8

Case Picture1.Width / 28 * 18 To Picture1.Width / 28 * 19
Beep (G1 * 4), Dur
PlaceNote Green, 6

Case Picture1.Width / 28 * 19 To Picture1.Width / 28 * 20
Beep (A1 * 4), Dur
PlaceNote Green, 4

Case Picture1.Width / 28 * 20 To Picture1.Width / 28 * 21
Beep (B1 * 4), Dur
PlaceNote Green, 2

Case Picture1.Width / 28 * 21 To Picture1.Width / 28 * 22
Beep (C1 * 8), Dur
PlaceNote Red, 13

Case Picture1.Width / 28 * 22 To Picture1.Width / 28 * 23
Beep (D1 * 8), Dur
PlaceNote Red, 11

Case Picture1.Width / 28 * 23 To Picture1.Width / 28 * 24
Beep (E1 * 8), Dur
PlaceNote Red, 9

Case Picture1.Width / 28 * 24 To Picture1.Width / 28 * 25
Beep (F1 * 8), Dur
PlaceNote Red, 8

Case Picture1.Width / 28 * 25 To Picture1.Width / 28 * 26
Beep (G1 * 8), Dur
PlaceNote Red, 6

Case Picture1.Width / 28 * 26 To Picture1.Width / 28 * 27
Beep (A1 * 8), Dur
PlaceNote Red, 4

Case Picture1.Width / 28 * 27 To Picture1.Width / 28 * 28
Beep (B1 * 8), Dur
PlaceNote Red, 2
End Select

Else
Select Case X

Case 240 To 495
Beep CS1, Dur
PlaceNote 0&, 12

Case 600 To 855
Beep DS1, Dur
PlaceNote 0&, 10

Case 1320 To 1575
Beep FS1, Dur
PlaceNote 0&, 7

Case 1680 To 1935
Beep GS1, Dur
PlaceNote 0&, 5

Case 2040 To 2295
Beep AS1, Dur
PlaceNote 0&, 3

Case 2760 To 3015
Beep (CS1 * 2), Dur
PlaceNote Blue, 12

Case 3120 To 3475
Beep (DS1 * 2), Dur
PlaceNote Blue, 10

Case 3840 To 4095
Beep (FS1 * 2), Dur
PlaceNote Blue, 7

Case 4200 To 4455
Beep (GS1 * 2), Dur
PlaceNote Blue, 5

Case 4560 To 4815
Beep (AS1 * 2), Dur
PlaceNote Blue, 3

Case 5280 To 5535
Beep (CS1 * 4), Dur
PlaceNote Green, 12

Case 5640 To 5895
Beep (DS1 * 4), Dur
PlaceNote Green, 10

Case 6360 To 6615
Beep (FS1 * 4), Dur
PlaceNote Green, 7

Case 6720 To 6975
Beep (GS1 * 4), Dur
PlaceNote Green, 5

Case 7080 To 7335
Beep (AS1 * 4), Dur
PlaceNote Green, 3

Case 7800 To 8055
Beep (CS1 * 8), Dur
PlaceNote Red, 12

Case 8160 To 8415
Beep (DS1 * 8), Dur
PlaceNote Red, 10

Case 8880 To 9135
Beep (FS1 * 8), Dur
PlaceNote Red, 7

Case 9240 To 9495
Beep (GS1 * 8), Dur
PlaceNote Red, 5

Case 9600 To 9855
Beep (AS1 * 8), Dur
PlaceNote Red, 3

End Select

End If
Picture2.SetFocus
End Sub

Private Function PlaceNote(Color As Variant, Div As Long)
Load Note(Note.UBound + 1)
Note(Note.UBound).Top = (Picture2.Height / 14 * Div) - Note(Note.UBound).Height / 2
StepStart = StepStart + StepF
Note(Note.UBound).Left = StepStart
Note(Note.UBound).FillColor = Color
Note(Note.UBound).Visible = True
End Function

Private Sub Picture2_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case Shift

Case 1
Select Case LCase(Chr(KeyCode))
 Case "a"
 Beep (A1 * 2), Tempo
 Case "b"
 Beep (B1 * 2), Tempo
 Case "c"
 Beep (C1 * 2), Tempo
 Case "d"
 Beep (D1 * 2), Tempo
 Case "e"
 Beep (E1 * 2), Tempo
 Case "f"
 Beep (F1 * 2), Tempo
 Case "g"
 Beep (G1 * 2), Tempo
 End Select
 
Case 0
Select Case LCase(Chr(KeyCode))
 Case "a"
 Beep A1, Tempo
 Case "b"
 Beep B1, Tempo
 Case "c"
 Beep C1, Tempo
 Case "d"
 Beep D1, Tempo
 Case "e"
 Beep E1, Tempo
 Case "f"
 Beep F1, Tempo
 Case "g"
 Beep G1, Tempo
 End Select
End Select

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox Y
End Sub

Public Function CalcFreq(Pos As Long, Color As Long)

Select Case Pos

 Case 1500 To 1530
    Select Case Color
    Case 0&
      CalcFreq = C1
    Case Blue
      CalcFreq = (C1 * 2)
    Case Green
      CalcFreq = (C1 * 4)
    Case Red
      CalcFreq = (C1 * 8)
    End Select
    
Case 1380 To 1410
Select Case Color
    Case 0&
      CalcFreq = CS1
    Case Blue
      CalcFreq = (CS1 * 2)
    Case Green
      CalcFreq = (CS1 * 4)
    Case Red
      CalcFreq = (CS1 * 8)
    End Select
    
Case 1260 To 1290
Select Case Color
    Case 0&
      CalcFreq = D1
    Case Blue
      CalcFreq = (D1 * 2)
    Case Green
      CalcFreq = (D1 * 4)
    Case Red
      CalcFreq = (D1 * 8)
    End Select
    
Case 1140 To 1170
Select Case Color
    Case 0&
      CalcFreq = DS1
    Case Blue
      CalcFreq = (DS1 * 2)
    Case Green
      CalcFreq = (DS1 * 4)
    Case Red
      CalcFreq = (DS1 * 8)
    End Select
    
Case 1035 To 1075
Select Case Color
    Case 0&
      CalcFreq = E1
    Case Blue
      CalcFreq = (E1 * 2)
    Case Green
      CalcFreq = (E1 * 4)
    Case Red
      CalcFreq = (E1 * 8)
    End Select
    
Case 915 To 945
Select Case Color
    Case 0&
      CalcFreq = F1
    Case Blue
      CalcFreq = (F1 * 2)
    Case Green
      CalcFreq = (F1 * 4)
    Case Red
      CalcFreq = (F1 * 8)
    End Select
    
Case 795 To 825
Select Case Color
    Case 0&
      CalcFreq = FS1
    Case Blue
      CalcFreq = (FS1 * 2)
    Case Green
      CalcFreq = (FS1 * 4)
    Case Red
      CalcFreq = (FS1 * 8)
    End Select
    
Case 675 To 705
Select Case Color
    Case 0&
      CalcFreq = G1
    Case Blue
      CalcFreq = (G1 * 2)
    Case Green
      CalcFreq = (G1 * 4)
    Case Red
      CalcFreq = (G1 * 8)
    End Select
    
Case 570 To 600
Select Case Color
    Case 0&
      CalcFreq = GS1
    Case Blue
      CalcFreq = (GS1 * 2)
    Case Green
      CalcFreq = (GS1 * 4)
    Case Red
      CalcFreq = (GS1 * 8)
    End Select
    
Case 450 To 480
Select Case Color
    Case 0&
      CalcFreq = A1
    Case Blue
      CalcFreq = (A1 * 2)
    Case Green
      CalcFreq = (A1 * 4)
    Case Red
      CalcFreq = (A1 * 8)
    End Select
    
Case 330 To 360
Select Case Color
    Case 0&
      CalcFreq = AS1
    Case Blue
      CalcFreq = (AS1 * 2)
    Case Green
      CalcFreq = (AS1 * 4)
    Case Red
      CalcFreq = (AS1 * 8)
    End Select
    
Case 210 To 240
Select Case Color
    Case 0&
      CalcFreq = B1
    Case Blue
      CalcFreq = (B1 * 2)
    Case Green
      CalcFreq = (B1 * 4)
    Case Red
      CalcFreq = (B1 * 8)
    End Select
    
Case 0 To 160
CalcFreq = P1
End Select

End Function

Private Sub text1_Change()




End Sub

Private Sub PlayTimer_Timer()
IntCnt = IntCnt + 1
Play.X1 = Play.X1 + StepF
Play.X2 = Play.X2 + StepF
Beep CalcFreq(Note(IntCnt).Top + (Note(IntCnt).Height / 2), Note(IntCnt).FillColor), Tempo
If IntCnt = Note.UBound Then PlayTimer.Enabled = False: IntCnt = 0: Picture2.SetFocus
End Sub

Private Sub Slider1_Change()
Tempo = Slider1.Value
PlayTimer.Interval = Tempo + 10
Picture2.SetFocus
End Sub

