VERSION 5.00
Begin VB.Form ViewFeedback 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ViewFeedback.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Last"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "First"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Label5"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Label5"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "ViewFeedback"
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
Label3.Caption = rs.Fields(2).Value
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
End Sub

Private Sub Command2_Click()

' Move previous
If rs.AbsolutePosition = 0 Then
    Command2.Enabled = False
Else
rs.MovePrevious
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Label3.Caption = rs.Fields(2).Value
Command1.Enabled = True
End If
End Sub

Private Sub Command3_Click()

' Move last


rs.MoveLast
Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Label3.Caption = rs.Fields(2).Value
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Form_Load()

' Form Load

Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from feedback")
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

Label7.Font.Size = 10
Label7.Font.Name = "arial"
Label7.Caption = "Rating"

Label4.Font.Size = 30
Label4.Font.Name = "arial"
Label4.Caption = "Customer Feedback"

'Label8.Font.Size = 10
'Label8.Font.Name = "arial"
'Label8.Caption = "Average Rating"

Label1.Caption = rs.Fields(0).Value
Label2.Caption = rs.Fields(1).Value
Label3.Caption = rs.Fields(2).Value

Command2.Enabled = False
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub

