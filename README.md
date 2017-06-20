# Information-Retrieval-Search-Engine
GUI: 1(Splash Screen 1)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Timer1.Enabled = True
    PBcolor ProgressBar1, vbRed, vbBlue
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value > 75 Then
        Label2.Caption = "Finalizing..."
        ElseIf ProgressBar1.Value > 57 Then
        Label2.Caption = "Gathering Information..."
        ElseIf ProgressBar1.Value > 33 Then
        Label2.Caption = "Initializing..."
    End If
    If ProgressBar1.Value > 99 Then
        Login.Show
        Timer1.Enabled = False
        Unload Me
    End If
End Sub


GUI :2(Login Form)

Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Image2_Click()
Registration.Show
Login.Hide
End Sub

Private Sub Image3_Click()
Dim X As Integer
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
While Not rs.EOF
If rs.Fields("Username").Value = Text1.Text Then
    If rs.Fields("Password").Value = Text2.Text Then
    If rs.Fields("Admin").Value = True Then
    adminForm.Show
    Module1.usr = Text1.Text
    rs.Close
    con.Close
    Unload Me
    Exit Sub
    End If
    MsgBox "Not Administrator"
    rs.Close
    con.Close
    Exit Sub
    End If
End If
rs.MoveNext
Wend
X = MsgBox("UNREGISTERED!", vbCritical)
rs.Close
con.Close
End Sub

Private Sub LoginMe_Click()
Dim user, pass As String
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
While Not rs.EOF
If rs.Fields("Username").Value = Text1.Text Then
    If rs.Fields("Password").Value = Text2.Text Then
    If rs.Fields("Status").Value = True Then
    Splash1.Show
    Else
    MsgBox "   Account" & vbCrLf & "Deactivated"
    End If
    rs.Close
    con.Close
    Unload Me
    Exit Sub
    End If
    MsgBox "Invalid Password"
    rs.Close
    con.Close
    Exit Sub
End If
rs.MoveNext
Wend
MsgBox ("fail")
rs.Close
con.Close
End Sub

GUI: 3(Registration Form)

Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Back_Click()
Login.Show
Unload Me
End Sub

Private Sub Image1_Click()
On Error Resume Next
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic

While Not rs.EOF
If rs.Fields("Username").Value = Username1.Text Then
    MsgBox "Username Used"
    rs.Close
    con.Close
    Exit Sub
End If
rs.MoveNext
Wend
rs.AddNew
rs.Fields("Name").Value = Name1.Text
rs.Fields("EmailId").Value = EmailID1.Text
rs.Fields("Username").Value = Username1.Text
rs.Fields("Password").Value = Password1.Text
rs.Fields("DOB").Value = DTPicker1.Value
rs.Update
MsgBox "            ADDED" & vbCrLf & " Activation Pending"
rs.Close
con.Close
Login.Show
Unload Me
End Sub

GUI: 4(Admin Portal)
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Login.Show
Unload Me
End Sub

Private Sub Command2_Click()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
While Not rs.EOF
    If Text2.Text = rs.Fields("Username").Value Then
    rs.Fields("Admin").Value = True
    MsgBox "Admin Added"
    rs.Update
    rs.Close
    con.Close
    Exit Sub
    End If
    rs.MoveNext
Wend
rs.Close
con.Close
End Sub

Private Sub Command3_Click()
addAdmin.Visible = False
Pending.Visible = False
Deactive.Visible = False
actionIdentify.Visible = True
End Sub

Private Sub Action_Click()
addAdmin.Visible = False
actionIdentify.Visible = False
Pending.Visible = False
Deactive.Visible = False
If (Op1.Value = True) Then
addAdmin.Visible = True
End If
If Op2.Value = True Then
Pending.Visible = True
End If
If Op3.Value = True Then
Deactive.Visible = True
End If
End Sub

Private Sub Command4_Click()
Splash1.Show
Unload Me
End Sub

Private Sub Command5_Click()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
While Not rs.EOF
    If Text2.Text = rs.Fields("Username").Value Then
    rs.Fields("Admin").Value = False
    MsgBox "Admin Removed"
    rs.Update
    rs.Close
    con.Close
    Exit Sub
    End If
    rs.MoveNext
Wend
rs.Close
con.Close
End Sub

Private Sub Deactivate_Click()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
While Not rs.EOF
    If Text6.Text = rs.Fields("Username").Value Then
    rs.Fields("Status").Value = False
    rs.Fields("Admin").Value = False
    MsgBox "Account Deactivated"
    rs.Update
    rs.Close
    con.Close
    Exit Sub
    End If
    rs.MoveNext
Wend
rs.Close
con.Close
End Sub

Private Sub Command6_Click()
'On Error Resume Next
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
While Not rs.EOF
    If rs.Fields("Username").Value = Text3.Text Then
        rs.Fields("Status").Value = True
        MsgBox "Activated"
        'Me.Refresh
        rs.Update
        rs.Close
        con.Close
        Exit Sub
    End If
    rs.MoveNext
Wend
rs.Close
con.Close
End Sub


Private Sub Dlist_Click()
Text6.Text = Dlist.Text
Dlist.clear
End Sub

Private Sub List1_Click()
Text3.Text = List1.Text
List1.Visible = False
End Sub

Private Sub Listsrc_Click()
Text2.Text = Listsrc.Text
Listsrc.Visible = False
End Sub

Private Sub Text1_Change()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
Listsrc.clear
While Not rs.EOF
    If UCase(Text1.Text) = UCase(Mid(rs!Name, 1, Len(Text1.Text))) Then
    Listsrc.AddItem rs.Fields("Username").Value
    End If
    rs.MoveNext
Wend
    If Listsrc.ListCount > 0 Then
    Listsrc.Visible = True
    Else
    Listsrc.Visible = False
    End If
rs.Close
con.Close
End Sub

Private Sub Text4_Change()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
List1.clear
While Not rs.EOF
    If UCase(Text4.Text) = UCase(Mid(rs!Name, 1, Len(Text4.Text))) Then
    List1.AddItem rs.Fields("Username").Value
    End If
    rs.MoveNext
Wend
If List1.ListCount > 0 Then
    List1.Visible = True
Else
    List1.Visible = False
End If
rs.Close
con.Close
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text5_Change()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
Dlist.clear
While Not rs.EOF
    If UCase(Text5.Text) = UCase(Mid(rs!Name, 1, Len(Text5.Text))) Then
    Dlist.AddItem rs.Fields("Username").Value
    End If
    rs.MoveNext
Wend
rs.Close
con.Close
End Sub

GUI: 5(Splash Screen 2)

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

GUI: 6(Selection Form)

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

GUI: 7(Search Form)

Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim i As Integer
Private Sub Command1_Click()
FoxitCtl1.Visible = False
List1.Visible = False
dataupdate (Text1.Text)
WebBrowser1.Navigate ("https://www.google.co.in/?gfe_rd=cr&ei=0e_8WM69Ee7I8AfrgouAAQ#q=" & Text1.Text)
End Sub
Private Sub Command2_Click()
List1.Visible = False
FoxitCtl1.Visible = False
dataupdate (Text1.Text)
WebBrowser1.Navigate ("https://in.search.yahoo.com/search;_ylc=X3oDMTFiN25laTRvBF9TAzIwMjM1MzgwNzUEaXRjAzEEc2VjA3NyY2hfcWEEc2xrA3NyY2h3ZWI-?p=" & Text1.Text & "&fr=yfp-t-704&fp=1&toggle=1&cop=mss&ei=UTF-8")
End Sub
Private Sub Command3_Click()
FoxitCtl1.Visible = False
List1.Visible = False
dataupdate (Text1.Text)
WebBrowser1.Navigate ("https://www.bing.com/search?q=" & Text1.Text & "&qs=n&form=QBRE&sp=-1&pq=" & Text1.Text & "&sc=8-5&sk=&cvid=AFFC4737AE6C4CD1A3E3A44475D08C40")
End Sub
Private Sub Command4_Click()
FoxitCtl1.Visible = False
List1.Visible = False
dataupdate (Text1.Text)
WebBrowser1.Navigate (Text1.Text)
End Sub
Private Sub Command5_Click()
con.Close
Choice.Show
Unload Me
End Sub
Private Sub Command6_Click()
con.Close
Login.Show
Unload Me
End Sub
Private Sub Form_Load()
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
    rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
    While Not (Module1.usr = rs!UserName)
        rs.MoveNext
    Wend
    listadd
End Sub
Sub listadd()
    List1.clear
    List1.AddItem rs.Fields("search3").Value
    List1.AddItem rs.Fields("search2").Value
    List1.AddItem rs.Fields("search1").Value
End Sub
Sub dataupdate(latestdata As String)
    If Not ((latestdata = rs.Fields("search2").Value) Or (latestdata = rs.Fields("search3").Value)) Then
        rs.Fields("search1").Value = rs.Fields("search2").Value
        rs.Fields("search2").Value = rs.Fields("search3").Value
        rs.Fields("search3").Value = latestdata
        rs.Update
    ElseIf (latestdata = rs.Fields("search2").Value) Then
        rs.Fields("search2").Value = rs.Fields("search3").Value
        rs.Fields("search3").Value = latestdata
    End If
End Sub

Private Sub List1_Click()
Text1.Text = List1.Text
List1.Visible = False
End Sub

Private Sub Text1_Change()
Dim str As String
On Error Resume Next
If (Text1.Text = "") Then
    List1.Visible = True
    listadd
  Else
    List1.Visible = False
End If
str = FindFile("b:", Text1.Text & ".pdf")
If (Right(str, Len(Text1.Text & ".pdf")) = (Text1.Text & ".pdf")) Then
    FoxitCtl1.OpenFile (str)
    FoxitCtl1.Visible = True
    dataupdate (Text1.Text)
End If
End Sub
Function FindFile(BasedirectoryName As String, FileName As String) As String
Dim filestring As String
Dim folderCol As New Collection
Dim folder
BasedirectoryName = Trim$(BasedirectoryName)
If Right$(BasedirectoryName, 1) <> "\" Then
  BasedirectoryName = BasedirectoryName & "\"
End If
If Dir$(BasedirectoryName & FileName, vbNormal Or vbArchive Or vbHidden Or vbSystem) <> "" Then
    FindFile = BasedirectoryName & FileName
    Exit Function
End If
filestring = Dir$(BasedirectoryName & "*", vbArchive Or vbHidden Or vbSystem Or vbDirectory)
Do While filestring <> ""
  DoEvents
On Error GoTo nxt
    If (GetAttr(BasedirectoryName & filestring) And vbDirectory) = vbDirectory Then
      If Left$(filestring, 1) <> "." And Left$(filestring, 2) <> ".." Then
      folderCol.Add BasedirectoryName & filestring & "\", BasedirectoryName & filestring & "\"
      End If
    End If
nxt:
  filestring = Dir$
  Trim$ (filestring)
Loop
For Each folder In folderCol
    FindFile = FindFile(CStr(folder), FileName)
    If FindFile <> "" Then Exit Function
Next folder
End Function
GUI: 8(Database Search Form)
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
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data,Source=B:\VB6\MYPROJECT\StudentPro.mdb;Persist Security Info=True"
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
con1.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\proj1.mdb;Persist Security Info=False"
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
Dim X, Y As String
List1.clear
rs.MoveFirst
While Not rs.EOF
Y = UCase(rs!Name)
X = UCase(Text6.Text)
If X = Mid(Y, 1, Len(Text6.Text)) Then
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
