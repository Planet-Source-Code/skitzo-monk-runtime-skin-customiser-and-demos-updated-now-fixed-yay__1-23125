VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm2 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.CommandButton CommandButton2 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   855
      VariousPropertyBits=   19
      Caption         =   "Load"
      Size            =   "1508;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   855
      VariousPropertyBits=   19
      Caption         =   "Exit"
      Size            =   "1508;873"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
On Error GoTo woops
Dim sfile As String
With CommonDialog1
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "Picture files (*.bmp;*.jpg;*.gif;*.bif)|*.bmp;*.jpg;*.gif;*.bif"
    .ShowOpen
    If Len(.fileName) = 0 Then Exit Sub
    sfile = .fileName
End With
frm2.Picture = LoadPicture(sfile)
woops:
End Sub

Private Sub Form_Load()
frm2.Picture = LoadPicture(App.path + "\bbTmpJpg.bmp")
End Sub
