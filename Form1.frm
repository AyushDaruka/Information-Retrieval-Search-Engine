VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   6120
   ClientTop       =   3465
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   9375
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      IMEMode         =   3  'DISABLE
      Left            =   4440
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2480
      Width           =   3360
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   425
      Left            =   4440
      TabIndex        =   4
      Top             =   1560
      Width           =   3360
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   7440
      Top             =   3480
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "Table1"
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
   Begin VB.Image Image3 
      Height          =   1800
      Left            =   360
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   3000
      Picture         =   "Form1.frx":5A40
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Image LoginMe 
      Height          =   615
      Left            =   2880
      Picture         =   "Form1.frx":770D
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   2475
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Vladimir Script"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Vladimir Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personalized Information Retrieval System"
      BeginProperty Font 
         Name            =   "Rage Italic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
   Begin VB.Image Image1 
      Height          =   4815
      Left            =   0
      Picture         =   "Form1.frx":E2DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim usg As String
Private Sub Image2_Click()
Registration.Show
Login.Hide
End Sub

Private Sub Image3_Click()
Dim x As Integer
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=B:\VB6\MY PROJECT\login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
While Not rs.EOF
If rs.Fields("Username").Value = Text1.Text Then
    If rs.Fields("Password").Value = Text2.Text Then
    If rs.Fields("Admin").Value = True Then
    adminForm.Show
    rs.Close
    con.Close
    Unload Me
    Exit Sub
    Else
    x = MsgBox("Not Administrator", vbCritical)
    rs.Close
    con.Close
    Exit Sub
    End If
    Else
    x = MsgBox("PASSWORD" & vbCrLf, vbCritical)
    rs.Close
    con.Close
    Exit Sub
    End If
End If

rs.MoveNext
Wend
x = MsgBox("UNREGISTERED!", vbCritical)
rs.Close
con.Close
End Sub

Private Sub LoginMe_Click()
Dim user, pass As String
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
rs.MoveFirst
While Not rs.EOF
If rs.Fields("Username").Value = Text1.Text Then
    If rs.Fields("Password").Value = Text2.Text Then
    If rs.Fields("Status").Value = True Then
    Splash1.Show
    usg = Text1.Text
    Module1.usr = usg
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

Login.Show
End Sub
