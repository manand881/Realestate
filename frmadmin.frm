VERSION 5.00
Begin VB.Form frmadmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmadmin.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Update"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add New"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Last"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "First"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label5"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim counter As Integer
Private Sub CancelButton_Click()

' Move First

rs.MoveFirst
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Command2.Enabled = False
Command1.Enabled = True
End Sub

Private Sub Command1_Click()

' Move Next

rs.MoveNext
On Error GoTo here2
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Label3.Caption = rs.Fields(2).Value
Command2.Enabled = True
If rs.AbsolutePosition = counter - 2 Then
here2:
Command1.Enabled = False
rs.MoveLast
End If
Command2.Enabled = True
End Sub

Private Sub Command2_Click()

' Move previous
If rs.AbsolutePosition = 0 Then
    Command2.Enabled = False
Else
rs.MovePrevious
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Command1.Enabled = True
End If
Command1.Enabled = True
End Sub

Private Sub Command3_Click()

' Move last


rs.MoveLast
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command4_Click()
If (Not rs.Fields(0).Value = "owner") Then
rs.Delete
MsgBox "Login Credentials Deleted!"
Else
MsgBox "Cannot delete owner login"
End If
End Sub

Private Sub Command5_Click()
'Add New
If Text1.Text = "" Or Text2.Text = "" Or Text1.Text = "Enter Id" Or Text2.Text = "Enter Password" Then

MsgBox "Enter new user ID and Password in the text box"

Else
rs.AddNew
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
Text1.Text = ""
Text2.Text = ""
rs.Update
MsgBox "New ID and Password has been added"
rs.MoveLast
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Command2.Enabled = True
End If
End Sub

Private Sub Command6_Click()

'Update

If Text1.Text = "" Or Text2.Text = "" Or Text1.Text = "Enter Id" Or Text2.Text = "Enter Password" Then

MsgBox "Enter new user ID and Password in the text box"

Else
rs.Edit
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
Text1.Text = ""
Text2.Text = ""
rs.Update
MsgBox "New ID and Password has been added"
rs.MoveLast
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Command2.Enabled = True
End If

End Sub

Private Sub Form_Load()

' Form Load

Text1.Text = "Enter Id"
Text2.Text = "Enter Password"

Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from Admin")
rs.MoveFirst
Sum = 0
counter = 1
While (rs.BOF = False)
counter = counter + 1
rs.MoveNext
On Error GoTo here
Wend
here:
rs.MoveFirst
'MsgBox (counter)

Label5.Font.Size = 10
Label5.Font.Name = "arial"
Label5.Caption = "Name"

Label6.Font.Size = 10
Label6.Font.Name = "arial"
Label6.Caption = "Phone Number"

Label4.Font.Size = 30
Label4.Font.Name = "arial"
Label4.Caption = "Admin Panel"

'Label8.Font.Size = 10
'Label8.Font.Name = "arial"
'Label8.Caption = "Average Rating"

Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value

Command2.Enabled = False
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub


