VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form signup 
   BackColor       =   &H00C0C000&
   Caption         =   "Form6"
   ClientHeight    =   10452
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17760
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   10452
   ScaleWidth      =   17760
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdclear1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   15000
      TabIndex        =   10
      Top             =   9360
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   5880
      Top             =   9600
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5525
      _ExtentY        =   1926
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\mesdc21\Documents\hys\nammadatabase.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\mesdc21\Documents\hys\nammadatabase.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "signin"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   2175
   End
   Begin VB.TextBox txtph 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   9000
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   6960
      Width           =   4095
   End
   Begin VB.TextBox txtem 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   9000
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox txtano 
      Height          =   735
      Left            =   9000
      MaxLength       =   12
      TabIndex        =   4
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txtname 
      Height          =   735
      Left            =   9000
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   7
      Top             =   7080
      Width           =   4335
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adhaar card number"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   3
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name                                            *as per aadhar"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sign up"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset


Private Sub cmdclear1_Click()
txtname.Text = ""
txtano.Text = ""
txtem.Text = ""
txtph.Text = ""
End Sub

Private Sub Command1_Click()
Dim str1 As String
If txtname = "" Then
MsgBox ("Enter name")
Else
If txtano = "" Then
MsgBox ("Enter aadhar number")
Else

If txtem = "" Then
MsgBox ("Enter Password")
Else
If txtph = "" Then

MsgBox ("Confirm Password")
Else
If txtem.Text = txtph.Text Then
MsgBox ("Password is matching")
End If

If txtem.Text <> txtph.Text Then
MsgBox ("Enter Correct Password")
Else
'search coding
rs.MoveFirst
ans = txtano.Text
rs.Find ("aadharno=" & ans)
If rs.EOF Or rs.BOF Then
rs.AddNew
rs(0) = txtname.Text
rs(1) = txtano.Text
rs(2) = txtem.Text
rs.Update
rs.Close
MsgBox ("Added new record")
Form3.Show
Else
MsgBox ("aadharno existing")
txtano.Text = ""
txtname.Text = ""
txtem.Text = ""
txtph.Text = ""
Exit Sub
End If
Form3.Show
str1 = "select * from nammalogin"

rs1.Open str1, cn, adOpenDynamic, adLockOptimistic

rs1.AddNew
rs1(0) = txtname.Text

rs1(1) = txtem.Text
rs1.Update
MsgBox ("Added new record to nammalogin")
End If
End If
End If
End If
End If

End Sub

Private Sub Command2_Click()
'search coding
rs.MoveFirst
ans = txtano.Text
rs.Find ("aadharno=" & ans)
If rs.EOF Or rs.BOF Then
rs.AddNew
rs(0) = txtname.Text
rs(1) = txtano.Text
rs(2) = txtem.Text
rs.Update
rs.Close
MsgBox ("Added new record")
Form3.Show
Else
MsgBox ("aadharno existing")
txtano.Text = ""
txtname.Text = ""
txtem.Text = ""
txtph.Text = ""
Exit Sub
End If
Form3.Show

End Sub

Private Sub Form_Load()
Dim str As String
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\HOME\Desktop\hys\nammadatabase.mdb;Persist Security Info=False"
str = "select * from signin"
cn.Open
rs.Open str, cn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub txtano_KeyPress(keyascii As Integer)
If keyascii >= 48 And keyascii <= 57 Then
Else
keyascii = 0
MsgBox "ENTER ONLY NUMBERS"
End If
End Sub

Private Sub txtem_keypress(keyascii As Integer)
If keyascii >= 48 And keyascii <= 57 Then
Else
keyascii = 0
MsgBox "ENTER ONLY NUMBERS"
End If
End Sub

Private Sub txtph_keypress(keyascii As Integer)
If keyascii >= 48 And keyascii <= 57 Then
Else
keyascii = 0
MsgBox "ENTER ONLY NUMBERS"
End If
End Sub
