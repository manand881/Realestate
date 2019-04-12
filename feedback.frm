VERSION 5.00
Begin VB.Form frmfeedback 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   2910
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "feedback.frx":0000
   Picture         =   "feedback.frx":11F3B
   ScaleHeight     =   2910
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "clear"
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "feedback.frx":23E76
      Left            =   2280
      List            =   "feedback.frx":23E89
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmfeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim textval As String
Dim numval As String
Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Load()

Text1.Text = ""
Label1.Font.Size = 10
Label1.Caption = "Name"
Label1.Font.Name = "arial"

Text2.Text = ""
Label2.Font.Size = 10
Label2.Caption = "Phone Number"
Label2.Font.Name = "arial"

Label3.Font.Size = 10
Label3.Caption = "Rating"
Label3.Font.Name = "arial"

Label4.Caption = "Feedback Form"
Label4.Font.Size = 30
Label4.Font.Name = "arial"

Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from feedback")
End Sub

Private Sub OKButton_Click()
If Text1.Text = "" Or Text2.Text = "" Or List1.Text = "" Then
    MsgBox "Invalid Entry, try again!"
    If List1.Text = "" Then
        MsgBox "Select List element by clicking on it and try again!"
    End If
    Text1.SetFocus
    
Else
        rs.AddNew
        rs.Fields(0).Value = Text1.Text
        rs.Fields(1).Value = Text2.Text
        rs.Fields(2).Value = List1.Text
        rs.Update
        MsgBox "Thank you for your valueable feedback!"
        Text1.Text = ""
        Text2.Text = ""
        List1.ListIndex = 0
End If
End Sub
Private Sub Text2_Change()
  textval = Text2.Text
  If IsNumeric(textval) Then
    numval = textval
  Else
    Text2.Text = ""
  End If
  
End Sub


