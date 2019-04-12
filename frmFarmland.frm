VERSION 5.00
Begin VB.Form frmSell 
   Caption         =   "Form2"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16200
   LinkTopic       =   "Form2"
   Picture         =   "frmFarmland.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   16200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   8640
      TabIndex        =   18
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Corner Site"
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   5760
      Width           =   2415
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
      Caption         =   "Delete"
      Height          =   615
      Left            =   8880
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Submit"
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
   Begin VB.Label Label7 
      Caption         =   "Commission"
      Height          =   615
      Left            =   6000
      TabIndex        =   17
      Top             =   1200
      Width           =   2415
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
Attribute VB_Name = "frmSell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim textval As String
Dim numval As String
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

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then

MsgBox "Invalid entry, try again!"

Else

rs.AddNew
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = Text4.Text
rs.Fields(4).Value = Text5.Text
rs.Fields(5).Value = Text6.Text

If Check1.Value = False Then

rs.Fields(6).Value = "No"

Else

rs.Fields(6).Value = " Yes"

End If
MsgBox "Your Details have been added to our database we will contact you when we find a buyer."
rs.Update
End If


End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()

'Form Load

Label1.Font.Size = 12
Label2.Font.Size = 12
Label3.Font.Size = 12
Label4.Font.Size = 12
Label5.Font.Size = 12
Label6.Font.Size = 12
Label7.Font.Size = 12
Check1.Font.Size = 12
Label1.Font.Name = "arial"
Label2.Font.Name = "arial"
Label3.Font.Name = "arial"
Label4.Font.Name = "arial"
Label5.Font.Name = "arial"
Label6.Font.Name = "arial"
Label7.Font.Name = "arial"
Check1.Font.Name = "arial"

Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from Seller")

End Sub

Private Sub Text3_Change()
 textval = Text3.Text
  If Isnumeric(textval) Then
    numval = textval
  Else
    Text3.Text = ""
  End If
End Sub

Private Sub Text4_Change()
 textval = Text4.Text
  If Isnumeric(textval) Then
    numval = textval
  Else
    Text4.Text = ""
  End If
  Text7.Text = Text4.Text * 0.005
End Sub

Private Sub Text6_Change()
  textval = Text6.Text
  If Isnumeric(textval) Then
    numval = textval
  Else
    Text6.Text = ""
  End If
  
  End Sub
Private Sub Text5_Change()
If (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Then
Else
keyascii = 0
Text5.Text = ""
End If
End Sub


