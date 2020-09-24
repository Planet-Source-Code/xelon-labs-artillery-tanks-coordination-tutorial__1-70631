VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artillery"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Controls"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   5535
      Begin VB.CommandButton Command7 
         Caption         =   "Save Scape"
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Load scape"
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Reimage"
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Teleport"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Fire"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Power 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Text            =   "500"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Angle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "90"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "New Space"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Power"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Angle"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0069332C&
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   471
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   2370
         Left            =   8880
         Pattern         =   "*.gif*"
         TabIndex        =   16
         Top             =   4560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog cdlg 
         Left            =   5520
         Top             =   5760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open landscape from GIF"
         Filter          =   "*.gif"
      End
      Begin VB.Timer tmraim 
         Enabled         =   0   'False
         Interval        =   6
         Left            =   1560
         Top             =   1680
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   9000
         ScaleHeight     =   225
         ScaleWidth      =   1215
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         Begin VB.Image Image2 
            Height          =   225
            Left            =   0
            Picture         =   "Form1.frx":17374
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Timer tmrexp 
         Enabled         =   0   'False
         Interval        =   6
         Left            =   1080
         Top             =   1680
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   600
         Top             =   1680
      End
      Begin VB.Image Image4 
         Height          =   180
         Left            =   7680
         Picture         =   "Form1.frx":181C8
         Top             =   4440
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Image3 
         Height          =   180
         Left            =   7080
         Picture         =   "Form1.frx":182B0
         Top             =   4200
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Damage = "
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   10200
         Picture         =   "Form1.frx":18398
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label wnd 
         BackStyle       =   0  'Transparent
         Caption         =   "Wind ="
         ForeColor       =   &H00F2E2D9&
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin VB.Image shape1 
         Height          =   750
         Left            =   5400
         Picture         =   "Form1.frx":191EC
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Image bomb 
         Height          =   165
         Left            =   8280
         Picture         =   "Form1.frx":19CDE
         Top             =   3960
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Image tgt 
         Height          =   1560
         Left            =   1680
         Picture         =   "Form1.frx":19EE7
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image command1 
         Height          =   300
         Left            =   8400
         Picture         =   "Form1.frx":1DB2B
         Top             =   5880
         Width           =   600
      End
   End
   Begin VB.Image img 
      Height          =   255
      Left            =   5640
      Top             =   7200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H0059341C&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":1DD17
      ForeColor       =   &H00F2E2D9&
      Height          =   1095
      Left            =   5760
      TabIndex        =   12
      Top             =   7200
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Dim w As Integer, n As Integer, pw As Integer, dmg As Integer, dn As Boolean
Dim mX As Single, mY As Single

Private Sub Angle_Change()
crsr
End Sub

Private Sub Command2_Click()
Timer1 = True
bomb.Visible = True
w = 1
n = 1
End Sub

Private Sub Command3_Click()
'This is the ability of teleporting
Dim X As Integer, Y As Integer, d As Boolean
Image3.Visible = False ' Eye target is hidden
e:
X = Rnd * Picture1.Width / 15   ' The left value is set to random
Y = Rnd * Picture1.Height / 15  ' and yes the top value is also random thats why the tank and target fly after teleporting
If GetPixel(Picture1.hdc, X, Y) <> Picture1.BackColor Then ' If the generated position is in land then try again
GoTo e
End If
If d = False Then
tgt.Left = X: tgt.Top = Y
d = True
GoTo e
Else
command1.Left = X: command1.Top = Y
End If
End Sub

Private Sub Command4_Click()
' Creating a new land scape
If dn = False Then dn = True: img.Picture = Picture1.Picture ' The origional image is stored in img
Set Picture1.Picture = Nothing ' Picture is cleared
Image3.Visible = False ' Eye target is hidden
Dim mdX1 As Long, mdX2 As Long, m1 As Integer, m2 As Integer, a1 As Long, a2 As Long
Dim rand As Integer, X As Integer, Y As Integer, sn As Integer
a1 = (Rnd * Picture1.ScaleWidth) / 2: a2 = ((Rnd * Picture1.ScaleWidth) / 2) + Picture1.ScaleWidth / 2 ' Random numbers are generated
mdX1 = a1 + (tgt.Width / 2): mdX2 = a2 + (command1.Width / 2)
While rand < Picture1.ScaleHeight / 2 ' The number is matched as if it is below the half of screen
rand = Rnd * (Picture1.Height / 15)     ' If not we try again
Wend
For X = 1 To (Picture1.Width / 15) ' loop is initiated to width of land scape
For Y = rand To Picture1.Height / 15 ' Rand contains the height from the bottom, This loop fills the land from height to depth
SetPixel Picture1.hdc, X, Y, &H8000& + (Y / (Picture1.ScaleHeight / 200)) ' land is made
Next
If X = mdX1 Then m1 = rand
If X = mdX2 Then m2 = rand
sn = Sgn(Rnd - Rnd) '' A new random number is generated 1 or -1
rand = rand - (sn * 2) ' if 1 then land is dunk else land is raised
Next
tgt.Left = a1: command1.Left = a2
tgt.Top = m1 - tgt.Height
command1.Top = m2 - command1.Height
End Sub

Private Sub Command5_Click()
' The image stored earlier in img is retrieved
If dn = True Then Picture1.Picture = img.Picture
End Sub

Private Sub Command6_Click()
cdlg.ShowOpen
Set Picture1.Picture = LoadPicture(cdlg.FileName)
End Sub

Private Sub Command7_Click()
cdlg.ShowSave
SavePicture Picture1.Picture, cdlg.FileName
End Sub

Private Sub File1_Click()
Set Picture1.Picture = LoadPicture(File1.Path & "\" & File1)
File1.Visible = False
End Sub

Private Sub Form_Load()
rebuild
End Sub

Private Sub Picture1_DblClick()
cdlg.ShowColor
Picture1.BackColor = cdlg.Color
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
'All key press events are here
If KeyCode = 38 Then
If Shift = 1 Then       '\
Angle = Angle - 8 '      '\
ElseIf Shift = 2 Then         '\
Angle = Angle - 5                '\
Else                                        ' angle value is changed according to Shift and arrow keys(/\-|-\/)
Angle = Angle - 1            '  /
End If                            '/
crsr
ElseIf KeyCode = 37 Then
command1.Left = command1.Left - 2
Image3.Visible = False
ElseIf KeyCode = 39 Then                    ' Tank is moved by arrow keys (<-|->)
Image3.Visible = False
command1.Left = command1.Left + 2
ElseIf KeyCode = 40 Then
If Shift = 1 Then
Angle = Angle + 3
ElseIf Shift = 2 Then
Angle = Angle + 2
Else
Angle = Angle + 1
End If
crsr
ElseIf KeyCode = 32 Then
tmraim = True
Image4.Visible = True
End If
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
tmraim = False
Image4.Visible = False
Command2_Click
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
File1.Visible = False
ElseIf Button = 2 Then
File1.Path = App.Path & "\BG\"
File1.Left = X: File1.Top = Y
File1.Visible = True
End If
End Sub

Private Sub tgt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False ' Eye target is hidden
If Button = 1 Then
mX = X
mY = Y
End If
End Sub

Private Sub tgt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
tgt.Move tgt.Left + (X - mX) / 15, tgt.Top + (Y - mY) / 15
End If
End Sub
Private Sub command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False ' Eye target is hidden
If Button = 1 Then
mX = X
mY = Y
End If
End Sub

Private Sub command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
command1.Move command1.Left + (X - mX) / 15, command1.Top + (Y - mY) / 15
End If
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
'Refer to Tutorial for more information
'The response of movement of Ball
w = w + 4 ' The distance value is increased
'Timer1.Interval = Timer1.Interval - 0.0001 ' Enable this for slight increase in speed
Dim mdX As Integer, mdY As Integer, pt As Integer, X As Integer 'As breifly explained in tutorial, The position of ballis set
bomb.Left = Cos(3.1416 * Angle / 180) * w - (w ^ 2) * pw / (50 * Power) + command1.Left
bomb.Top = -(Sin(3.1416 * Angle / 180) * w) + (w ^ 2) / Power + command1.Top
If GetPixel(Picture1.hdc, bomb.Left, bomb.Top) <> Picture1.BackColor Then
task ' The ball collides with the land
End If
If bomb.Left > tgt.Left And bomb.Left < tgt.Left + tgt.Width Then     ' The ball collides with with target
If bomb.Top > tgt.Top And bomb.Top < tgt.Top + tgt.Height Then
mdX = (tgt.Left + tgt.Left + tgt.Width) / 2
mdY = (tgt.Top + tgt.Top + tgt.Height) / 2
pt = Abs((bomb.Width * bomb.Height) - (Abs(bomb.Left - mdX) * Abs(bomb.Top - mdY)) / 2) / 15 'Hit Points are found
dmg = dmg + pt ' Damage is increased according to points, For mare damage try to hit on sides of target
If dmg > 70 And dmg < 80 Then MsgBox "Yo did the damage, You won. Read Tutorial for more details", vbDefaultButton3, "You won"
Label1 = "Damage = " & dmg ' Damage is displayed
task
End If
End If
Exit Sub
End Sub

Sub task() ' If the ball collides with tower or land
'Every thing is packed up
Timer1 = False
bomb.Visible = False
shape1.Visible = True
shape1.Height = 1
shape1.Width = 1
shape1.Top = bomb.Top
shape1.Left = bomb.Left
Power = 1
tmrexp = True ' The exploding timer is turned on
End Sub

Private Sub tmraim_Timer()
Power = Val(Power) + 10
Scrsr
End Sub

Private Sub tmrexp_Timer()
'Exploding effect after hitting
' The explode image is enlarged continously
shape1.Left = shape1.Left - 1
shape1.Width = shape1.Width + 2
shape1.Top = shape1.Top - 1
shape1.Height = shape1.Height + 2
If n > 60 Then ' 60 is diameter of the blast, you can change it
shape1.Visible = False
tmrexp = False
destroy bomb.Left, bomb.Top  ' This will destroy the land from x and y
rebuild
End If
n = n + 3
End Sub

Sub rebuild()
' Every thing is redone
Dim rd As Single
rd = Rnd
pw = rd * 30
pw = pw * Sgn(Rnd - Rnd)
Picture2.Visible = False
If Sgn(pw) = 1 Then
Picture2.Visible = True
Image1.Visible = False
Picture2.Width = rd * 80
Picture2.Left = 680 - Picture2.Width
Image2.Left = (Picture2.Width * 15) - Image2.Width
ElseIf Sgn(pw) = 0 Then
Image1.Visible = False
Picture2.Visible = False
Else
Picture2.Visible = False
Image1.Visible = True
Image1.Width = rd * 80
End If
wnd = "Wind = " & pw
End Sub

Sub destroy(X As Integer, Y As Integer)
On Error Resume Next
' This Sub is used to destroy land from X and Y
Dim a As Long, b As Long, c As Long, d As Long
shape1.Visible = False
For b = 0 To 30 Step 1 ' Change 30 to increased value for big hole, Also try Step 3, 7, 10, 20
For a = 0 To 360 Step 4 ' Remove Step 4 for total blast, Also try Step 3, 7, 10, 20
c = CLng(Cos(3.1416 * a / 180) * b + X)
d = CLng(-(Sin(3.1416 * a / 180) * b) + Y)
Picture1.PSet (c, d), Picture1.BackColor
Next: Next
End Sub

Sub crsr()
' The cursor of angle when we use Up and Down keys
Image3.Visible = True
Image3.Left = Cos(3.1416 * Angle / 180) * 40 + command1.Left   ' It is placed according to the angle
Image3.Top = -(Sin(3.1416 * Angle / 180) * 40) + command1.Top
End Sub
Sub Scrsr()
' When we hold spacebar another cursor appears
Image4.Left = Cos(3.1416 * Angle / 180) * Power / 15 + command1.Left ' The distance of cursor from tank is the power
Image4.Top = -(Sin(3.1416 * Angle / 180) * Power / 15) + command1.Top
End Sub
