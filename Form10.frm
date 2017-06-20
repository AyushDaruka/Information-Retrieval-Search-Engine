VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form StudentDatabase 
   ClientHeight    =   7980
   ClientLeft      =   6120
   ClientTop       =   2460
   ClientWidth     =   10590
   LinkTopic       =   "Form3"
   ScaleHeight     =   7980
   ScaleWidth      =   10590
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7935
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   10575
      ExtentX         =   18653
      ExtentY         =   13996
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1320
      Left            =   6360
      TabIndex        =   19
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Text            =   "Search by name"
      Top             =   120
      Width           =   3975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "Find"
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   8
      Top             =   3160
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   7
      Top             =   4480
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3120
      TabIndex        =   6
      Top             =   5160
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   5
      Text            =   "SEMESTER"
      Top             =   3840
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   159776769
      CurrentDate     =   42819
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "F"
      DisabledPicture =   "Form10.frx":0000
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "M"
      DisabledPicture =   "Form10.frx":2A2C
      Height          =   255
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
      Width           =   400
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      Picture         =   "Form10.frx":5458
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1155
   End
   Begin VB.Image AddImage 
      Height          =   855
      Left            =   8400
      Picture         =   "Form10.frx":B8AE
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Update 
      Height          =   615
      Left            =   6120
      Picture         =   "Form10.frx":214FF
      Stretch         =   -1  'True
      Top             =   6950
      Width           =   2295
   End
   Begin VB.Image Delete 
      Height          =   735
      Left            =   2595
      Picture         =   "Form10.frx":281CD
      Stretch         =   -1  'True
      Top             =   6900
      Width           =   2415
   End
   Begin VB.Image Add 
      Height          =   1080
      Left            =   4920
      Picture         =   "Form10.frx":2D716
      Stretch         =   -1  'True
      Top             =   6700
      Width           =   1200
   End
   Begin VB.Image Pervious 
      Height          =   735
      Left            =   0
      Picture         =   "Form10.frx":323DE
      Stretch         =   -1  'True
      Top             =   6880
      Width           =   2655
   End
   Begin VB.Image Next 
      Height          =   750
      Left            =   8400
      Picture         =   "Form10.frx":3A0AB
      Stretch         =   -1  'True
      Top             =   6900
      Width           =   2160
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   4540
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   3200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ROLL No.:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   1320
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   8655
      Left            =   0
      Picture         =   "Form10.frx":419AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "StudentDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim con1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim str As String

Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo1.Text = ""
DTPicker1.Value = Date - 7000
Option1.Value = Option2.Value = False
End Sub

Private Sub AddImage_Click()
'Dim x As String
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Image1.Picture = LoadPicture(str)
End Sub

Private Sub Command3_Click()
rs.MoveFirst
If Val(Text1.Text) > 0 Then
roll
Else
Name1
End If
con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\proj1.mdb;Persist Security Info=False"
rs1.Open "Select * from proj", con1, adOpenDynamic, adLockReadOnly
rs1.MoveFirst
While Not rs1.EOF
    If Text1.Text = rs1.Fields("Roll").Value Then
        List2.AddItem rs1!Project
    End If
    rs1.MoveNext
Wend
rs1.Close
con1.Close
End Sub
Sub Name1()
While Not rs.EOF
If UCase((Mid(rs!Name, 1, Len(Text2.Text))) = UCase(Text2.Text)) Then
display
Exit Sub
End If
rs.MoveNext
Wend
End Sub
Sub roll()
While Not rs.EOF
If rs.Fields("Enroll").Value = Val(Text1.Text) Then
display
Exit Sub
End If
rs.MoveNext
Wend
End Sub

Sub refreshdata()
rs.Close
rs.Open "Select *from StudPro", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
display
Else
MsgBox "No Records"
End If
End Sub

Private Sub Delete_Click()
confirm = MsgBox("Do you want to delete?", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rs.Delete adAffectCurrent
rs.Update
refreshdata
Else
MsgBox "Not deleted"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\StudentPro.mdb;Persist Security Info=True"
rs.Open "Select * from StudPro", con, adOpenDynamic, adLockPessimistic
rs.Sort = "Roll ASC"
rs.Close
rs.Open "Select * from StudPro", con, adOpenDynamic, adLockPessimistic

    Combo1.AddItem ("SEMESTER I")
    Combo1.AddItem ("SEMESTER II")
    Combo1.AddItem ("SEMESTER III")
    Combo1.AddItem ("SEMESTER IV")
    Combo1.AddItem ("SEMESTER V")
    Combo1.AddItem ("SEMESTER VI")
    Combo1.AddItem ("SEMESTER VII")
    Combo1.AddItem ("SEMESTER VIII")
End Sub

Private Sub Add_Click()
rs.AddNew
clear
List1.Visible = False
End Sub


Private Sub Image3_Click()
If (WebBrowser1.Visible = True) Then
    WebBrowser1.Visible = False
    Text6.Visible = True
  Else
Choice.Show
Unload Me
End If
End Sub

Private Sub List1_Click()
Text2.Text = List1.Text
List1.Visible = False
Text6.Text = ""
End Sub

Private Sub List2_Click()
con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=proj1.mdb;Persist Security Info=False"
rs1.Open "Select * from proj", con1, adOpenDynamic, adLockPessimistic
rs1.MoveFirst
While Not rs1.EOF
    If (List2.Text = rs1!Project) Then
        WebBrowser1.Navigate (rs1!Link)
        WebBrowser1.Visible = True
        Text6.Visible = False
    End If
    rs1.MoveNext
Wend
rs1.Close
con1.Close
End Sub

Private Sub Next_Click()
On Error Resume Next
rs.MoveNext
If rs.EOF Then
rs.MoveFirst
End If
display
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
Option2.Value = False
Text5.Text = "Male"
End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Option1.Value = False
Text5.Text = "Female"
End If
End Sub
Sub display()
Text1.Text = rs!Enroll
Text2.Text = rs!Name
DTPicker1.Value = rs!DOB
Text4.Text = rs!Result
Image1.Picture = LoadPicture(rs!Image)
Combo1.Text = rs!Semester
Text5.Text = rs!Gender
Text3.Text = rs!Contact

End Sub

Private Sub Pervious_Click()
On Error Resume Next
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
End If
display
End Sub

Private Sub Save_Click()
rs.Fields("Name").Value = Text2.Text
rs.Fields("Enroll").Value = Text1.Text
If Option1.Value = True Then
    rs.Fields("Gender").Value = Option1.Caption
   ElseIf Option2.Value = True Then
    rs.Fields("Gender").Value = Option2.Caption
End If
rs.Fields("Semester").Value = Combo1.Text
rs.Fields("Result").Value = Text4.Text
rs.Fields("Contact").Value = Text3.Text
rs.Fields("Image").Value = str
rs.Fields("DOB").Value = DTPicker1.Value
rs.Update
MsgBox "Data added Successfully"
StudentDatabase.Show

End Sub

Private Sub Text6_Change()
Dim x, y As String
List1.clear
rs.MoveFirst
While Not rs.EOF
y = UCase(rs!Name)
x = UCase(Text6.Text)
If x = Mid(y, 1, Len(Text6.Text)) Then
    List1.AddItem rs!Name
End If
rs.MoveNext
Wend
If ((List1.ListCount > 0) And (Text6.Text <> "")) Then
List1.Visible = True
Else
List1.Visible = False
End If
End Sub

Private Sub Text6_Click()
Text6.Text = ""
Text6.Forecolor = vbBlack
End Sub

Private Sub Update_Click()
On Error Resume Next
rs.Fields("Name").Value = Text2.Text
rs.Fields("Enroll").Value = Text1.Text
If Option1.Value = True Then
    rs.Fields("Gender").Value = Option1.Caption
   ElseIf Option2.Value = True Then
    rs.Fields("Gender").Value = Option2.Caption
End If
rs.Fields("Semester").Value = Combo1.Text
rs.Fields("Result").Value = Text4.Text
rs.Fields("Contact").Value = Text3.Text
rs.Fields("Image").Value = str
rs.Fields("DOB").Value = DTPicker1.Value
rs.Update
MsgBox "Data added Successfully"
StudentDatabase.Show
End Sub

