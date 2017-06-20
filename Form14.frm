VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Registration 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   4515
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   5280
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.TextBox Password1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   425
      Left            =   2760
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Username1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   425
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox EmailID1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   425
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox Name1 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   425
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   425
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   -2147483646
      CalendarTitleBackColor=   -2147483635
      Format          =   112656385
      CurrentDate     =   42819
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION  FORM"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   3000
      Picture         =   "Form14.frx":B41F
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Image Back 
      Height          =   495
      Left            =   0
      Picture         =   "Form14.frx":12228
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DOB:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EMAIL ID:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Back_Click()
Login.Show
Unload Me
End Sub

Private Sub Image1_Click()
On Error Resume Next
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
