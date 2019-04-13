VERSION 5.00
Begin VB.Form frmBuy 
   Caption         =   "Form2"
   ClientHeight    =   9435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   Picture         =   "frmBuy.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   615
      Left            =   2880
      TabIndex        =   24
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Move Last"
      Height          =   615
      Left            =   10320
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Next"
      Height          =   615
      Left            =   8880
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Previous"
      Height          =   615
      Left            =   7440
      TabIndex        =   20
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Move First"
      Height          =   615
      Left            =   6000
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "I Agree to the terms and conditions"
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   6720
      Width           =   2415
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   8880
      TabIndex        =   17
      Top             =   1920
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   615
      Left            =   10320
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Terms"
      Height          =   615
      Left            =   8880
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buy"
      Height          =   615
      Left            =   7440
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   615
      Left            =   6000
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   4800
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   2880
      TabIndex        =   9
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "Corner Site"
      Height          =   615
      Left            =   240
      TabIndex        =   23
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Commission"
      Height          =   615
      Left            =   6000
      TabIndex        =   16
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Phone Number"
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Name"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Asking Price"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "SqFeet"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Locality"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "City"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmBuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Public db_profit As Database
Public rs_profit As Recordset
Dim textval As String
Dim numval As String
Dim counter As Integer

Private Sub Command1_Click()
'clear all

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Check1.Value = False

End Sub

Private Sub Command2_Click()
'Move First

rs.MoveFirst
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(6).Value

Command7.Enabled = True
Command6.Enabled = False

End Sub

Private Sub Command3_Click()

'buy
If Check2.Value = False Then
MsgBox "You cannot submit the form untill you accept our terms and conditions"
GoTo skip
End If
rs_profit.AddNew
rs_profit.Fields(0) = Text7.Text * 2
rs_profit.Update
rs.Delete

MsgBox "congratulations on your new property"
skip:
End Sub

Private Sub Command4_Click()
Terms.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()

'Move Previous
Command7.Enabled = True
rs.MovePrevious
If rs.AbsolutePosition = 0 Then
Command6.Enabled = False
Else
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(6).Value
End If
End Sub

Private Sub Command7_Click()

'Move Next
If rs.AbsolutePosition = counter Then
Command7.Enabled = False
Else
rs.MoveNext
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(6).Value
Command6.Enabled = True
End If
End Sub

Private Sub Command8_Click()

'Move Last

rs.MoveLast
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(6).Value
Command7.Enabled = False
Command6.Enabled = True
End Sub

Private Sub Form_Load()
'Form Load
WindowState = vbMaximized

Label1.Font.Size = 12
Label2.Font.Size = 12
Label3.Font.Size = 12
Label4.Font.Size = 12
Label5.Font.Size = 12
Label6.Font.Size = 12
Label7.Font.Size = 12
'Check1.Font.Size = 12
Label1.Font.Name = "arial"
Label2.Font.Name = "arial"
Label3.Font.Name = "arial"
Label4.Font.Name = "arial"
Label5.Font.Name = "arial"
Label6.Font.Name = "arial"
Label7.Font.Name = "arial"
'Check1.Font.Name = "arial"

Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from Seller")
Set db_profit = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs_profit = db.OpenRecordset("select * from Profit")

rs.MoveFirst
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
Text6.Text = rs.Fields(5).Value
Text7.Text = rs.Fields(7).Value
Text8.Text = rs.Fields(6).Value
Command6.Enabled = False

While (Not rs.EOF)
counter = rs.AbsolutePosition
rs.MoveNext
Wend
rs.MoveFirst
End Sub

Private Sub Text3_Change()
 textval = Text3.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    Text3.Text = ""
  End If
End Sub

Private Sub Text4_Change()
 textval = Text4.Text
  If IsNumeric(textval) Then
    numval = textval
    Text7.Text = Text4.Text * 0.005
  Else
    Text4.Text = ""
  End If
  
End Sub

Private Sub Text6_Change()
  textval = Text6.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    Text6.Text = ""
  End If
  
  End Sub
Private Sub Text5_keypress(keyascii As Integer)
If keyascii >= 65 And keyascii <= 122 Then
    If keyascii >= 91 And keyascii <= 96 Then
        Text5.Text = ""
    End If
Else
keyascii = 0
Text5.Text = ""
End If
End Sub


