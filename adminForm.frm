VERSION 5.00
Begin VB.Form adminForm 
   BackColor       =   &H00400000&
   ClientHeight    =   5655
   ClientLeft      =   6555
   ClientTop       =   2400
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   7815
   Begin VB.Frame actionIdentify 
      Caption         =   "CHOOSE ACTION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   7575
      Begin VB.OptionButton Op1 
         Caption         =   "Add or Remove Admin"
         BeginProperty Font 
            Name            =   "Gill Sans MT Condensed"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Action 
         Caption         =   "Action"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4320
         TabIndex        =   18
         Top             =   3600
         Width           =   3135
      End
      Begin VB.OptionButton Op2 
         Caption         =   "Activation Pending"
         BeginProperty Font 
            Name            =   "Gill Sans MT Condensed"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   3135
      End
      Begin VB.OptionButton Op3 
         Caption         =   "Deactivate Account"
         BeginProperty Font 
            Name            =   "Gill Sans MT Condensed"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   2655
      End
   End
   Begin VB.Frame Deactive 
      Caption         =   "Deactivate"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7575
      Begin VB.ListBox Dlist 
         BackColor       =   &H00C0C0C0&
         Height          =   1815
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   6855
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   360
         TabIndex        =   23
         Top             =   3645
         Width           =   3735
      End
      Begin VB.CommandButton Deactivate 
         Caption         =   "Deactivate"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4200
         TabIndex        =   22
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2520
         TabIndex        =   20
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label Label5 
         Caption         =   "Selected Username"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Pending 
      Caption         =   "Pending Activation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Text4 
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
         Height          =   435
         Left            =   240
         TabIndex        =   13
         Text            =   "Enter name"
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Approve"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         TabIndex        =   11
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   10
         Top             =   3120
         Width           =   5295
      End
      Begin VB.ListBox List1 
         BackColor       =   &H8000000A&
         Height          =   1815
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Label Label3 
         Caption         =   "User:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   2295
   End
   Begin VB.Frame addAdmin 
      Caption         =   "Admin Membership Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.ListBox Listsrc 
         BackColor       =   &H00C0C0C0&
         Height          =   645
         Left            =   3000
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Remove Admin"
         Height          =   615
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3600
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Make Admin"
         Height          =   615
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3600
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3120
         TabIndex        =   2
         Top             =   2760
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3000
         TabIndex        =   1
         Top             =   780
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER USERNAME:"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER NAME:"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "GET IN"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   28
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Caption         =   "LOG OUT"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   27
      Top             =   4920
      Width           =   2415
   End
End
Attribute VB_Name = "adminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Login.Show
Unload Me
End Sub

Private Sub Command2_Click()

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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

con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
While Not rs.EOF
    If rs.Fields("Username").Value = Text3.Text Then
        rs.Fields("Status").Value = True
        rs.Fields("search1").Value = "x"
        rs.Fields("search2").Value = "x"
        rs.Fields("search3").Value = "x"
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
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=login.mdb;Persist Security Info=False"
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
