VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Apresentação"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   240
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   240
      Top             =   240
   End
   Begin VB.Image Image3 
      Height          =   1185
      Left            =   2520
      Picture         =   "Form1.frx":030A
      Top             =   7440
      Visible         =   0   'False
      Width           =   7035
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   4800
      Picture         =   "Form1.frx":2C8B
      Top             =   7440
      Width           =   2490
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5745
      Left            =   840
      Picture         =   "Form1.frx":ABC1
      Top             =   960
      Width           =   10290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'res
End Sub

Private Sub Timer1_Timer()
    Image2.Visible = False
    Image3.Visible = True
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Unload Form10
    Unload Form11
    Unload Form12
    Unload Form15
    Form2.Visible = True
    Form1.Visible = False
    Timer2.Enabled = False
End Sub
