VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1500
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   886.25
   ScaleMode       =   0  'User
   ScaleWidth      =   3492.879
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   240
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2400
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset
Public n As Integer
Public password As String
Public username As String
Option Explicit
Public login As Boolean
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Unload Me
    Unload Main
    MsgBox "Goodbye!"
End Sub

Private Sub cmdOK_Click()
    rs.MoveFirst
    LoginSucceeded = False
    
    While (Not rs.EOF)
    n = rs.RecordCount
        If txtUserName = rs.Fields(0).Value And txtPassword = rs.Fields(1).Value Then
            LoginSucceeded = True
            Main.Show
            Unload Me
        End If
    rs.MoveNext
    Wend
    
    If (LoginSucceeded = False) Then
        MsgBox "Invalid Credentials, try again!", , "Login"
        txtPassword.SetFocus
    End If

End Sub

Private Sub Form_Load()
Set db = OpenDatabase("C:\Program Files\Microsoft Visual Studio\VB98\Realestate\test.mdb")
Set rs = db.OpenRecordset("select * from Admin")
rs.MoveFirst
End Sub
