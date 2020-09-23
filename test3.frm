VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form test3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Psyco Skin 3"
   ClientHeight    =   1335
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1905
   Icon            =   "test1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   1905
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ToggleButton ToggleButton2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   840
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
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
      Height          =   780
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2000
      ForeColor       =   255
      VariousPropertyBits=   8388627
      Caption         =   "Skin Test !!"
      Size            =   "3528;1376"
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
    .Filter = "Skitzo Skins for Test 3 (*.skz3)|*.skz3"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    sfile = .FileName
End With
test3.Picture = LoadPicture(sfile)
woops:
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
