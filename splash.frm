VERSION 5.00
Begin VB.Form splash 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   Picture         =   "splash.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Continue"
      Height          =   495
      Left            =   19080
      TabIndex        =   1
      Top             =   10200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "exit"
      Height          =   495
      Left            =   21120
      TabIndex        =   0
      Top             =   10200
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Height          =   915
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   12585
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "splash"
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
WindowState = vbMaximized
End Sub

