VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form test3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Psyco Skin 3"
   ClientHeight    =   4335
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4905
   HelpContextID   =   5000
   Icon            =   "test1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.SpinButton SpinButton1 
      Height          =   1335
      Left            =   4080
      TabIndex        =   7
      Top             =   2160
      Width           =   375
      Size            =   "661;2355"
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
      VariousPropertyBits=   746604563
      DisplayStyle    =   3
      Size            =   "3625;450"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.OptionButton OptionButton1 
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1815
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   5
      Size            =   "3201;450"
      Value           =   "0"
      Caption         =   "Radio Button"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3413;450"
      Value           =   "0"
      Caption         =   "CheckBox "
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
      VariousPropertyBits=   -1400879085
      Size            =   "6165;2355"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ToggleButton ToggleButton2 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3840
      Width           =   855
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "Logo on"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ToggleButton ToggleButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   855
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1508;661"
      Value           =   "0"
      Caption         =   "Menu on"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label1 
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4155
      ForeColor       =   255
      VariousPropertyBits=   8388627
      Caption         =   "Skin Test !!"
      Size            =   "7329;741"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuload 
         Caption         =   "Load Skin"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "test3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "Skin Test !!"
Label1.ForeColor = vbRed
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.Caption = "By Psyco Softwarez"
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuload_Click()
On Error GoTo woops
Dim sfile As String
With CommonDialog1
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "Skitzo Skins for Test 1 (*.skz1)|*.skz1"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    sfile = .FileName
End With
test3.Picture = LoadPicture(sfile)
woops:
End Sub

Private Sub SpinButton1_Change()
If TextBox1.TextAlign = fmTextAlignLeft Then
TextBox1.TextAlign = fmTextAlignCenter
Else
If TextBox1.TextAlign = fmTextAlignCenter Then
TextBox1.TextAlign = fmTextAlignRight
Else
If TextBox1.TextAlign = fmTextAlignRight Then
TextBox1.TextAlign = fmTextAlignLeft
End If
End If
End If
End Sub



Private Sub ToggleButton1_Click()
If ToggleButton1.Value = True Then
mnufile.Visible = False
Else
mnufile.Visible = True
End If
End Sub

Private Sub ToggleButton2_Click()
If ToggleButton2.Value = True Then
Label1.Visible = False
Else
Label1.Visible = True
End If
End Sub
