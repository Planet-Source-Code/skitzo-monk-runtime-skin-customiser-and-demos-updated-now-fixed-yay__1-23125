VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "For What Program"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Test3"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Test2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Test1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
MsgBox "It is done"
paint.MainPic.Height = 5000
paint.MainPic.Width = 5000
Unload Me
Else
If Option2.Value = True Then
MsgBox "It shall be done"
paint.MainPic.Height = 3000
paint.MainPic.Width = 3000
Unload Me

Else
If Option3.Value = True Then
MsgBox "It has been done"
paint.MainPic.Height = 2000
paint.MainPic.Width = 2000
Unload Me
End If
End If
End If
paint.Show
End Sub

