VERSION 5.00
Begin VB.Form Terms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
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
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Terms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Label1.Caption = "Terms & Conditions are as follows for sellers:" & vbNewLine & "" & vbNewLine & "-> You affirm that the declared details are true." & vbNewLine & "-> You agree to pay the company commission as shown." & vbNewLine & "-> You will not ask for black money." & vbNewLine & "-> We will not settle any disputes between you and the customer after the sale of your property" & vbNewLine & "" & vbNewLine & "Terms & Conditions are as follows for buyers:" & vbNewLine & "" & vbNewLine & "-> You will not offer to pay the seller in black money" & vbNewLine & "-> You agree to pay us our commission for our serviecs once you buy the property." & vbNewLine & "-> We will not settle any disputes between you and the seller "
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
