VERSION 5.00
Begin VB.Form Choice 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   7065
   ClientTop       =   2400
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   Picture         =   "Choice.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   8055
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "LOG OUT"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   4680
      Picture         =   "Choice.frx":B41F
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   840
      Picture         =   "Choice.frx":1251A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Login.Show
Unload Me
End Sub

Private Sub Image1_Click()
SearchBox.Show
Unload Me
End Sub

Private Sub Image2_Click()
StudentDatabase.Show
Unload Me
End Sub
