VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{3A8BD65E-9922-4162-A649-83F2D5326BBE}#1.0#0"; "FoxitReaderBrowserAx.dll"
Begin VB.Form SearchBox 
   BackColor       =   &H00800000&
   Caption         =   "SearchBox"
   ClientHeight    =   10200
   ClientLeft      =   5055
   ClientTop       =   1410
   ClientWidth     =   12210
   FillColor       =   &H00800000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   12210
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12240
      Top             =   1440
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   5055
   End
   Begin FOXITREADERLibCtl.FoxitCtl FoxitCtl1 
      Height          =   9255
      Left            =   100
      TabIndex        =   8
      Top             =   600
      Width           =   12015
      _cx             =   5080
      _cy             =   5080
      src             =   ""
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "LOG  OUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF0000&
      Caption         =   "URL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "BiNG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF0000&
      Caption         =   "YAHOO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   93
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "GOOGLE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   12015
      ExtentX         =   21193
      ExtentY         =   16325
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
      Location        =   ""
   End
End
Attribute VB_Name = "SearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
