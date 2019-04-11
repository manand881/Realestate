VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11040
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   Picture         =   "splash.frx":0000
   ScaleHeight     =   11040
   ScaleWidth      =   16080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   20040
      TabIndex        =   1
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   495
      Left            =   21600
      TabIndex        =   0
      Top             =   10320
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   915
      Left            =   465
      TabIndex        =   2
      Top             =   1080
      Width           =   12585
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
frmLogin.Show
Unload Me
End Sub

Private Sub Form_Load()
Label1.Font.Size = 32
Label1.Font.Name = ariel
Label1.Caption = "Welcome to the Automated Real-Estate Management System"
End Sub

