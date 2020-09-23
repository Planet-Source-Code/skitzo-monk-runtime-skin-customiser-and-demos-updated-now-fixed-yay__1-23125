VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form paint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Psyco Softwarez Skin Making Demo"
   ClientHeight    =   7305
   ClientLeft      =   1860
   ClientTop       =   1455
   ClientWidth     =   11625
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11625
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1111
      ButtonWidth     =   1402
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Line"
            Object.ToolTipText     =   "Draw a straight line!"
            Object.Tag             =   ""
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pencil"
            Object.ToolTipText     =   "Draw as you want!"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Eraser"
            Object.ToolTipText     =   "Erase the bad things!"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Box"
            Object.ToolTipText     =   "Draw boxes you want!"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Circle"
            Object.ToolTipText     =   "Draw circles you want!"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Picker"
            Object.ToolTipText     =   "Pick color you want from the picture!"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Caption         =   "Paint"
            Object.ToolTipText     =   "Paint the area you want! But Not available now."
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Gradation"
            Object.ToolTipText     =   "Make the Gradation you want!"
            Object.Tag             =   ""
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   55
      Top             =   7200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   6960
      Top             =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   6000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CommandButton DialogBox 
      Caption         =   "SecondColor Dialog"
      Height          =   495
      Index           =   1
      Left            =   7440
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
   End
   Begin VB.HScrollBar Scroll 
      Height          =   135
      Index           =   5
      Left            =   6000
      Max             =   255
      TabIndex        =   15
      Top             =   5880
      Value           =   255
      Width           =   3015
   End
   Begin VB.HScrollBar Scroll 
      Height          =   135
      Index           =   4
      Left            =   6000
      Max             =   255
      TabIndex        =   14
      Top             =   5640
      Value           =   255
      Width           =   3015
   End
   Begin VB.HScrollBar Scroll 
      Height          =   135
      Index           =   3
      Left            =   6000
      Max             =   255
      TabIndex        =   13
      Top             =   5400
      Value           =   255
      Width           =   3015
   End
   Begin VB.HScrollBar Scroll 
      Height          =   135
      Index           =   2
      Left            =   6000
      Max             =   255
      TabIndex        =   11
      Top             =   4800
      Width           =   3015
   End
   Begin VB.HScrollBar Scroll 
      Height          =   135
      Index           =   1
      Left            =   6000
      Max             =   255
      TabIndex        =   10
      Top             =   4560
      Width           =   3015
   End
   Begin VB.HScrollBar Scroll 
      Height          =   135
      Index           =   0
      Left            =   6000
      Max             =   255
      TabIndex        =   9
      Top             =   4320
      Width           =   3015
   End
   Begin VB.CommandButton DialogBox 
      Caption         =   "FrontColor Dialog"
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox ForeColorSample 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   5
      ToolTipText     =   "FrontColor"
      Top             =   3000
      Width           =   615
   End
   Begin VB.PictureBox BackColorSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6240
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   6
      ToolTipText     =   "Back Color"
      Top             =   3240
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "JPEG files|*.jpg|GIF files|*.GIF|Bitmap files|*.BMP|All Files|*.*"
   End
   Begin VB.PictureBox ColorBoard 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6000
      MousePointer    =   2  'Cross
      ScaleHeight     =   1695
      ScaleWidth      =   3000
      TabIndex        =   1
      ToolTipText     =   "Left Click - FrontColor, Right Click - BackColor"
      Top             =   960
      Width           =   3000
   End
   Begin VB.PictureBox MainPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   6315
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   720
      Width           =   5895
      Begin VB.Timer tmrTimer 
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer Timer2 
         Left            =   840
         Top             =   1320
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3120
         Top             =   2160
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   1095
         Left            =   720
         Top             =   2640
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   1680
         X2              =   3720
         Y1              =   960
         Y2              =   2760
      End
   End
   Begin VB.Frame Optionframe 
      Caption         =   "Circle or Picker or Paint Option"
      Height          =   6255
      Index           =   3
      Left            =   9120
      TabIndex        =   44
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label13 
         Caption         =   "***Picker : Click the place (Left Click will pick frontcolor, Right Click will pick backcolor.)"
         Height          =   735
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Optionframe 
      Caption         =   "Gradation"
      Height          =   6255
      Index           =   4
      Left            =   9120
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Frame Frame1 
         Caption         =   "Direction"
         Height          =   975
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   2175
         Begin VB.OptionButton GradationDirection 
            Caption         =   "Horizontal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton GradationDirection 
            Caption         =   "Vertical"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Color"
         Height          =   1335
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   2175
         Begin VB.OptionButton GradationColor 
            Caption         =   "ForeColor to BackColor"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton GradationColor 
            Caption         =   "ForeColor to White"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton GradationColor 
            Caption         =   "ForeColor to Black"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000009&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   2955
         ScaleWidth      =   2115
         TabIndex        =   46
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Sample"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Frame Optionframe 
      Caption         =   "Box Option"
      Height          =   6255
      Index           =   1
      Left            =   9120
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Frame BoxOptionFrame 
         Caption         =   "Select Color of the Interior"
         Height          =   1335
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   2175
         Begin VB.OptionButton BoxOptionInterior 
            Caption         =   "White"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   39
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton BoxOptionInterior 
            Caption         =   "ForeColor"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton BoxOptionInterior 
            Caption         =   "BackColor"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton BoxOptionInterior 
            Caption         =   "Transparent"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.PictureBox BoxOptionPicture 
         BackColor       =   &H00E0E0E0&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   2115
         TabIndex        =   22
         Top             =   240
         Width           =   2175
         Begin VB.Shape BoxOptionSample 
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   600
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame Optionframe 
      Caption         =   "Line and Pencil Option"
      Height          =   6255
      Index           =   0
      Left            =   9120
      TabIndex        =   27
      Top             =   720
      Width           =   2415
      Begin VB.Frame LineOptionFrame 
         Caption         =   "Border Size"
         Height          =   1575
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton Command2 
            Caption         =   "Change"
            Height          =   375
            Left            =   720
            TabIndex        =   40
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox LineOptionText 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "1"
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Input the integer value between 1~10"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Input the Border Size"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1575
         End
         Begin VB.Line Line2 
            X1              =   1920
            X2              =   1920
            Y1              =   480
            Y2              =   1440
         End
      End
   End
   Begin VB.Frame Optionframe 
      Caption         =   "Eraser Option"
      Height          =   6255
      Index           =   2
      Left            =   9120
      TabIndex        =   32
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Frame EraserOptionFrame2 
         Caption         =   "Color of the erased place"
         Height          =   855
         Left            =   120
         TabIndex        =   41
         Top             =   3360
         Width           =   2175
         Begin VB.OptionButton EraserOptionColor 
            Caption         =   "White"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton EraserOptionColor 
            Caption         =   "BackColor"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame EraserOptionFrame 
         Caption         =   "Eraser Size"
         Height          =   3015
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton Command1 
            Caption         =   "Change"
            Height          =   375
            Left            =   1320
            TabIndex        =   38
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox EraserOptionText 
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   34
            Text            =   "300"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label10 
            Caption         =   "Sample"
            Height          =   255
            Left            =   840
            TabIndex        =   37
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Input the integer value between 100 ~ 500"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Input the Width of the eraser you want"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1575
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  'Opaque
            Height          =   300
            Left            =   960
            Top             =   2160
            Width           =   300
         End
      End
   End
   Begin VB.Label RGBValue 
      Caption         =   "RGB (255 , 255 , 255)"
      Height          =   255
      Index           =   1
      Left            =   7200
      TabIndex        =   18
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label RGBValue 
      Caption         =   "RGB (0 , 0 , 0)"
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   17
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "BackColor"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "ForeColor"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Pick the Color you want"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "Now the colors are.."
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Color Pick Board"
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      Begin VB.Menu SubMenuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu SubMenuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu SubMenuSaveAs 
         Caption         =   "Sa&ve As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu SubMenuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnutest 
      Caption         =   "Test"
   End
   Begin VB.Menu MenuFilter 
      Caption         =   "Fil&ter"
      Begin VB.Menu SubMenuBlur 
         Caption         =   "Mosaic"
      End
   End
End
Attribute VB_Name = "paint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EraserColor As Long
Dim EraserSize As Integer
Dim PencilSize As Integer
Dim BoxInversed As Boolean
Dim RndNum As Byte
Dim GradationChanged As Boolean
Dim XX As Double, YY As Double
Dim XX2 As Double, YY2 As Double
Dim CurrentChoice
Dim TheColor As Long
Dim Red As Long
Dim Green As Long
Dim Blue As Long
Dim SecondColor As Long
Dim FirstColor As Long
Private Sub BoxOptionInterior_Click(Index As Integer)
BoxOptionSample.BackStyle = IIf(Index = 2, 0, 1)
If Index = 0 Then BoxOptionSample.BackColor = FirstColor
If Index = 1 Then BoxOptionSample.BackColor = SecondColor
If Index = 3 Then BoxOptionSample.BackColor = &HFFFFFF
End Sub

Private Sub ColorBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo 10
TheColor = ColorBoard.Point(x, y)
If Button <> 1 And Button <> 2 Then Exit Sub
If Button = 1 Then ForeColorSample.BackColor = TheColor: FirstColor = TheColor: g = 0
If Button = 2 Then BackColorSample.BackColor = TheColor: SecondColor = TheColor: g = 3
Scroll(g).Value = TakeRGB(TheColor, 0): Scroll(g + 1).Value = TakeRGB(TheColor, 1): Scroll(g + 2).Value = TakeRGB(TheColor, 2)
10 End Sub
Private Sub Command1_Click()
f$ = InputBox("Input the size of the eraser", "Drawer V1.0", EraserOptionText.Text)
f$ = RTrim$(LTrim$(f$))
If " " + f$ <> Str$(Val(f$)) Then MsgBox "Input error!", vbOKOnly, "Drawer V1.0": Exit Sub
If Val(f$) <> Int(Val(f$)) Then MsgBox "Input error!", vbOKOnly, "Drawer V1.0": Exit Sub
If Val(f$) > 500 Or Val(f$) < 100 Then MsgBox "Input error!", vbOKOnly, "Drawer V1.0": Exit Sub
EraserOptionText.Text = f$
EraserSize = Val(f$)
Shape3.Width = Val(f$): Shape3.Height = Val(f$)
Shape1.Width = Val(f$): Shape1.Height = Val(f$)
End Sub
Private Sub Command2_Click()
f$ = InputBox("Input the border of the line or pencil", "Drawer V1.0", LineOptionText.Text)
f$ = RTrim$(LTrim$(f$))
If " " + f$ <> Str$(Val(f$)) Then MsgBox "Input error!", vbOKOnly, "Drawer V1.0": Exit Sub
If Val(f$) <> Int(Val(f$)) Then MsgBox "Input error!", vbOKOnly, "Drawer V1.0": Exit Sub
If Val(f$) > 10 Or Val(f$) < 1 Then MsgBox "Input error!", vbOKOnly, "Drawer V1.0": Exit Sub
LineOptionText.Text = f$
PencilSize = Val(f$)
Line2.borderwidth = Val(f$)
End Sub
Private Sub DialogBox_Click(Index As Integer)
Static coloring As Long
On Error GoTo 100
CommonDialog1.ShowColor
coloring = CommonDialog1.Color
Scroll(Index * 3).Value = TakeRGB(coloring, 0)
Scroll(Index * 3 + 1).Value = TakeRGB(coloring, 1)
Scroll(Index * 3 + 2).Value = TakeRGB(coloring, 2)
100
End Sub

Private Sub EraserOptionColor_Click(Index As Integer)
EraserColor = IIf(Index = 0, SecondColor, &HFFFFFF)
End Sub
Private Sub EraserOptionText_GotFocus()
Command1.SetFocus
End Sub
Private Sub Form_Load()
EraserColor = &HFFFFFF
PencilSize = 1
EraserSize = 300
CurrentChoice = 1
FirstColor = &H0
SecondColor = &HFFFFFF
MsgBox "this program was taken off the planet source code website - it was modiified to add a better opening and saving system and skinning capabilities added, The Credit for his work is well appreciated. Thnx"
MsgBox "This program requires microsoft forms 2.0 object library, Sorry if ya dont have it i dunno wer to get it!"
paint.Hide
Form1.Show
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Shape1.Visible = False
End Sub
Private Sub GradationColor_Click(Index As Integer)
GradationChanged = True
End Sub
Private Sub GradationDirection_Click(Index As Integer)
GradationChanged = True
End Sub
Private Sub LineOptionText_GotFocus()
Command2.SetFocus
End Sub



Private Sub MainPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Select Case CurrentChoice
    Case 1
        Line1.X1 = x: Line1.X2 = x
        Line1.Y1 = y: Line1.Y2 = y
        Line1.Visible = True
    Case 2
        XX = x: YY = y
    Case 3
        MainPic.Line (Shape1.Left, Shape1.Top)-(Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Width), EraserColor, BF
    Case 4, 5, 8
        XX = x: YY = y
        XX2 = x: YY2 = y
        Shape2.Shape = IIf(CurrentChoice = 5, 2, 0)
        Shape2.Visible = True
        Shape2.Left = x: Shape2.Top = y
        Shape2.Width = 0: Shape2.Height = 0
End Select
End Sub
Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If CurrentChoice = 3 Then
        Shape1.Left = x - Shape1.Width / 2
        Shape1.Top = y - Shape1.Width / 2
        Shape1.Visible = True
End If
If Button <> 1 Then GoTo 10
Select Case CurrentChoice
    Case 1
        Line1.X2 = x: Line1.Y2 = y
    Case 2
        MainPic.DrawWidth = PencilSize
        MainPic.Line (XX, YY)-(x, y), FirstColor: XX = x: YY = y
        MainPic.DrawWidth = 1
    Case 3
        MainPic.Line (Shape1.Left, Shape1.Top)-(Shape1.Left + Shape1.Width, Shape1.Top + Shape1.Width), EraserColor, BF
    Case 4, 5, 8
        XX2 = x: YY2 = y
        Shape2.Left = IIf(x > XX, XX, x)
        Shape2.Top = IIf(y > YY, YY, y)
        Shape2.Width = Abs(x - XX)
        Shape2.Height = Abs(y - YY)
    Case 6
        Scroll(0).Value = TakeRGB(MainPic.Point(x, y), 0)
        Scroll(1).Value = TakeRGB(MainPic.Point(x, y), 1)
        Scroll(2).Value = TakeRGB(MainPic.Point(x, y), 2)
End Select
Exit Sub
10 If Button <> 2 Or CurrentChoice <> 6 Then Exit Sub
Scroll(3).Value = TakeRGB(MainPic.Point(x, y), 0)
Scroll(4).Value = TakeRGB(MainPic.Point(x, y), 1)
Scroll(5).Value = TakeRGB(MainPic.Point(x, y), 2)
End Sub
Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
Select Case CurrentChoice
    Case 1
        MainPic.DrawWidth = PencilSize
        MainPic.Line (Line1.X1, Line1.Y1)-(Line1.X2, Line1.Y2), FirstColor
        MainPic.DrawWidth = 1
        Line1.Visible = False
    Case 4
        If BoxOptionInterior(0).Value = True Then MainPic.Line (XX, YY)-(XX2, YY2), FirstColor, BF
        If BoxOptionInterior(1).Value = True Then MainPic.Line (XX, YY)-(XX2, YY2), SecondColor, BF
        If BoxOptionInterior(3).Value = True Then MainPic.Line (XX, YY)-(XX2, YY2), &HFFFFFF, BF
        MainPic.Line (XX, YY)-(XX2, YY2), FirstColor, B
        Shape2.Visible = False
    Case 5
        Rad = IIf(Abs(YY2 - YY) > Abs(XX2 - XX), Abs(YY2 - YY) / 2, Abs(XX2 - XX) / 2)
        If XX2 <> XX Then MainPic.Circle ((XX2 + XX) / 2, (YY2 + YY) / 2), Rad, FirstColor, , , Abs(YY2 - YY) / Abs(XX2 - XX)
        Shape2.Visible = False
    Case 8
        Dim sc1 As Long
        Dim sc2 As Long
        sc1 = FirstColor
        If GradationColor(0).Value = True Then sc2 = SecondColor
        If GradationColor(1).Value = True Then sc2 = &HFFFFFF
        If GradationColor(2).Value = True Then sc2 = &H0
        f1 = TakeRGB(sc2, 0): f2 = TakeRGB(sc2, 1): f3 = TakeRGB(sc2, 2)
        v1 = TakeRGB(sc1, 0): v2 = TakeRGB(sc1, 1): v3 = TakeRGB(sc1, 2)
        forstep = 10
        If XX2 < XX Then xx3 = XX: XX = XX2: XX2 = xx3
        If YY2 < YY Then yy3 = YY: YY = YY2: YY2 = yy3
        ForStart = IIf(GradationDirection(0).Value = True, XX, YY)
        Endpro = IIf(GradationDirection(0).Value = True, XX2, YY2)
        For i = ForStart To Endpro Step forstep
        D1 = v1 + (f1 - v1) / (Endpro - ForStart) * (i - ForStart)
        D2 = v2 + (f2 - v2) / (Endpro - ForStart) * (i - ForStart)
        D3 = v3 + (f3 - v3) / (Endpro - ForStart) * (i - ForStart)
        If GradationDirection(0).Value = True Then MainPic.Line (i, YY)-(i, YY2), RGB(D1, D2, D3)
        If GradationDirection(1).Value = True Then MainPic.Line (XX, i)-(XX2, i), RGB(D1, D2, D3)
        Next i
        Shape2.Visible = False
End Select
End Sub

Private Sub mnuinv_Click()
Call Invert

End Sub
Private Function Invert()
Call BitBlt(MainPic.hdc, 0, 0, MainPic.Width, MainPic.Height, Pic.hdc, 0, 0, SRCINVERT)
Call MainPic_MouseDown(1, 1, -33, -33)
Call MainPic_MouseUp(1, 1, -33, -33)
End Function

Private Sub mnutest_Click()
Randomize Timer
sfile = App.path + "\bbTmpJpg.bmp"
MsgBox "converting"
SavePicture MainPic.Image, sfile
frm2.Show
End Sub

Private Sub Scroll_Change(Index As Integer)
P = Int(Index / 3)
RGBValue(P).Caption = "RGB (" + RTrim$(Str$(Scroll(P * 3).Value)) + " , " + RTrim$(Str$(Scroll(P * 3 + 1).Value)) + " , " + RTrim$(Str$(Scroll(P * 3 + 2).Value)) + " )"
TheColor = RGB(Scroll(P * 3).Value, Scroll(P * 3 + 1).Value, Scroll(P * 3 + 2).Value)
If P = 0 Then FirstColor = TheColor: ForeColorSample.BackColor = TheColor Else SecondColor = TheColor: BackColorSample.BackColor = TheColor
Line2.BorderColor = FirstColor
BoxOptionSample.BorderColor = FirstColor
If BoxOptionInterior(0).Value = True Then BoxOptionSample.BackColor = FirstColor
If BoxOptionInterior(1).Value = True Then BoxOptionSample.BackColor = SecondColor
GradationChanged = True
End Sub
Private Sub Scroll_Scroll(Index As Integer)
P = Int(Index / 3)
RGBValue(P).Caption = "RGB (" + RTrim$(Str$(Scroll(P * 3).Value)) + " , " + RTrim$(Str$(Scroll(P * 3 + 1).Value)) + " , " + RTrim$(Str$(Scroll(P * 3 + 2).Value)) + " )"
TheColor = RGB(Scroll(P * 3).Value, Scroll(P * 3 + 1).Value, Scroll(P * 3 + 2).Value)
If P = 0 Then FirstColor = TheColor: ForeColorSample.BackColor = TheColor Else SecondColor = TheColor: BackColorSample.BackColor = TheColor
Line2.BorderColor = FirstColor
BoxOptionSample.BorderColor = FirstColor
If BoxOptionInterior(0).Value = True Then BoxOptionSample.BackColor = FirstColor
If BoxOptionInterior(1).Value = True Then BoxOptionSample.BackColor = SecondColor
GradationChanged = True
End Sub
Function TakeRGB(Colors As Long, Index As Integer) As Long
IndexColor = Colors
Red = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Red) / 256
Green = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Green) / 256
Blue = IndexColor
If Index = 0 Then TakeRGB = Red
If Index = 1 Then TakeRGB = Green
If Index = 2 Then TakeRGB = Blue
End Function
Private Sub SubMenuBlur_Click()
f = 97: f2 = f / 2 - 1
All = (MainPic.ScaleWidth - f) * (MainPic.ScaleHeight - f) / f / f
For i = f2 To MainPic.ScaleWidth - f2 Step f
For j = f2 To MainPic.ScaleHeight - f2 Step f
R = 0: g = 0: b = 0
For k = -f2 To f2 Step f2 / 2: For l = -f2 To f2 Step f2 / 2
R = R + TakeRGB(MainPic.Point(i + k, j + l), 0)
g = g + TakeRGB(MainPic.Point(i + k, j + l), 1)
b = b + TakeRGB(MainPic.Point(i + k, j + l), 2)
Next l, k
MainPic.Line (i - f2, j - f2)-(i + f2, j + f2), RGB(R / 25, g / 25, b / 25), BF
h = h + 1
If h > All Then ProgressBar1.Value = 100 Else ProgressBar1.Value = h / All * 100
Next j
Next i
MsgBox "done!!!"
ProgressBar1.Value = 0
End Sub
Private Sub SubMenuExit_Click()
Unload Me
End Sub
Private Sub SubMenuNew_Click()
paint.MainPic.Cls
Form1.Show
End Sub
Private Sub SubMenuOpen_Click()
On Error GoTo woops
Dim sfile As String
With CommonDialog1
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "Picture files (*.bmp;*.jpg;*.gif;*.bif)|*.bmp;*.jpg;*.gif;*.bif|SkitzoSkins for Test 1 (*.skz1)|*.skz1|SkitzoSkins for Test 2 (*.skz2)|*.skz2|SkitzoSkins for Test 3 (*.skz3)|*.skz3"
    .ShowOpen
    If Len(.fileName) = 0 Then Exit Sub
    sfile = .fileName
    If .FilterIndex = 2 Then
    MainPic.Width = 5000
    MainPic.Height = 5000
    End If
    If .FilterIndex = 3 Then
    MainPic.Width = 3000
    MainPic.Height = 3000
    End If
    If .FilterIndex = 4 Then
    MainPic.Width = 2000
    MainPic.Height = 2000
    End If
End With
MainPic.Picture = LoadPicture(sfile)
woops:
End Sub


Private Sub SubMenuSaveAs_Click()
On Error GoTo woops
If MainPic.Width = 5000 Then
 With CommonDialog1
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "SkitzoSkins for test 1 (*.skz1)|*.skz1"
        .ShowSave
        If Len(.fileName) = 0 Then Exit Sub
        sfile = .fileName
        SavePicture MainPic.Image, sfile
End With
Else
If MainPic.Width = 3000 Then
    Dim rfile As String
    With CommonDialog1
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "SkitzoSkins for Test 2 (*.skz2)|*.skz2"
        .ShowSave
        If Len(.fileName) = 0 Then Exit Sub
        rfile = .fileName
        SavePicture MainPic.Image, rfile
End With
Else
If MainPic.Width = 2000 Then

    Dim tfile As String
    With CommonDialog1
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "SkitzoSkins for Test 3 (*.skz3)|*.skz3"
        .ShowSave
        If Len(.fileName) = 0 Then Exit Sub
        tfile = .fileName
        SavePicture MainPic.Image, tfile
End With



woops:
End If
End If
End If
End Sub

Private Sub timer1_timer()
If GradationChanged = False Then Exit Sub
Dim sc1 As Long
Dim sc2 As Long

sc1 = FirstColor
If GradationColor(0).Value = True Then sc2 = SecondColor
If GradationColor(1).Value = True Then sc2 = &HFFFFFF
If GradationColor(2).Value = True Then sc2 = &H0
f1 = TakeRGB(sc2, 0): f2 = TakeRGB(sc2, 1): f3 = TakeRGB(sc2, 2)
v1 = TakeRGB(sc1, 0): v2 = TakeRGB(sc1, 1): v3 = TakeRGB(sc1, 2)
ForStart = 0: forstep = 10
Endpro = IIf(GradationDirection(0).Value = True, Picture1.ScaleWidth, Picture1.ScaleHeight)
For i = ForStart To Endpro Step forstep
D1 = v1 + (f1 - v1) / Endpro * i
D2 = v2 + (f2 - v2) / Endpro * i
D3 = v3 + (f3 - v3) / Endpro * i
If GradationDirection(0).Value = True Then Picture1.Line (i, 0)-(i, Picture1.ScaleHeight), RGB(D1, D2, D3)
If GradationDirection(1).Value = True Then Picture1.Line (0, i)-(Picture1.ScaleWidth, i), RGB(D1, D2, D3)
10 Next i
GradationChanged = False
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
For i = 1 To 8
If Toolbar1.Buttons(i).Value = tbrPressed Then CurrentChoice = i
Next i
Shape1.Visible = False
Line1.Visible = False
For i = 0 To 4
Optionframe(i).Visible = False
Next i
Select Case CurrentChoice
    Case 1 To 2
        Optionframe(0).Visible = True
    Case 3
        Optionframe(2).Visible = True
    Case 4
        Optionframe(1).Visible = True
    Case 5 To 7
        Optionframe(3).Visible = True
    Case 8
        GradationChanged = True
        Optionframe(4).Visible = True
End Select
End Sub



