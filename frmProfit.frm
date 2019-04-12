VERSION 5.00
Begin VB.Form frmProfit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1980
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs_profit As Recordset
Public db1 As Database
Public rs_seller As Recordset
Dim Counter As Integer
Option Explicit

Private Sub Form_Load()
Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs_profit = db.OpenRecordset("select * from Profit")
'rs_profit.MoveFirst

Set db1 = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs_seller = db.OpenRecordset("select * from Seller")
rs_seller.MoveFirst

Counter = 0
While (Not rs_seller.EOF)
rs_profit.Fields(0).Value = rs_seller.Fields(7).Value
rs_profit.Update
rs_profit.MoveNext
rs_seller.MoveNext
Counter = Counter + 1
Wend

End Sub

