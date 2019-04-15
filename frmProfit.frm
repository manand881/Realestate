VERSION 5.00
Begin VB.Form frmProfit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1155
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Dim counter As Integer
Dim buffer As Double
Dim profit As Double
Option Explicit

Private Sub Form_Load()

Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from Profit")
rs.MoveFirst

counter = 0
While (Not rs.EOF)
buffer = rs.Fields(0).Value
profit = profit + buffer
rs.MoveNext
Wend

Label1.Caption = "Total Profit made till date is: " & profit & " Rupees"

End Sub


Private Sub OKButton_Click()
Unload Me
End Sub
