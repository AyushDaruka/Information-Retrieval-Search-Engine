VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   4515
   ClientLeft      =   6555
   ClientTop       =   2400
   ClientWidth     =   8850
   LinkTopic       =   "Form2"
   ScaleHeight     =   4515
   ScaleWidth      =   8850
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   3960
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   8280
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   50
      TabIndex        =   1
      Top             =   30
      Width           =   8775
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Personalized Information Retrieval System"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   3
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   2
         Top             =   3600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Splash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This section is for the API declares
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'This section is for Progressbar's
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)

'This section is for the Colors of the ProgressBar's
Public Sub PBcolor(PB As ProgressBar, Backcolor As Long, Forecolor As Long)
SendMessage PB.hwnd, CCM_SETBKCOLOR, 0, ByVal Backcolor
SendMessage PB.hwnd, PBM_SETBARCOLOR, 0, ByVal Forecolor
End Sub

Private Sub Form_Load()
Label3.Caption = "Personalized  Information" & vbCrLf & vbTab & vbTab & "                         Retrieval System"
Timer1.Enabled = True
PBcolor Bar1, vbBlue, &H8000000F
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Bar1.Value = Bar1.Value + 1
Label1.Caption = Bar1.Value & "%"
    If Bar1.Value > 75 Then
        Label2.Caption = "Finalizing..."
        ElseIf Bar1.Value > 57 Then
        Label2.Caption = "Gathering Information..."
        ElseIf Bar1.Value > 33 Then
        Label2.Caption = "Initializing..."
    End If
    If Bar1.Value > 99 Then
        Timer1.Enabled = False
        Label1.Caption = ""
        Choice.Show
        Unload Me
    End If
End Sub



