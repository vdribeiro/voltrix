VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Curiosidades"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form4"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      MaskColor       =   &H00000000&
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1935
      Left            =   7200
      MaskColor       =   &H00000000&
      Picture         =   "Form4.frx":CCA5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1935
      Left            =   4200
      MaskColor       =   &H00000000&
      Picture         =   "Form4.frx":D7F5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   1935
      Left            =   840
      MaskColor       =   &H00000000&
      Picture         =   "Form4.frx":E1EE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   555
      Left            =   1440
      TabIndex        =   7
      Top             =   8160
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ideias"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   9345
      TabIndex        =   5
      Top             =   6480
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dicas"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   6360
      TabIndex        =   4
      Top             =   3960
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Raças"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7680
      Index           =   0
      Left            =   0
      Picture         =   "Form4.frx":12605
      Top             =   0
      Width           =   7680
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7680
      Index           =   3
      Left            =   7680
      Picture         =   "Form4.frx":1525D
      Top             =   7680
      Width           =   7680
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7680
      Index           =   2
      Left            =   0
      Picture         =   "Form4.frx":17EB5
      Top             =   7680
      Width           =   7680
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   7680
      Index           =   1
      Left            =   7680
      Picture         =   "Form4.frx":1AB0D
      Top             =   0
      Width           =   7680
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form6.Visible = True
    Form4.Visible = False
End Sub

Private Sub Command2_Click()
    Form7.Visible = True
    Form4.Visible = False
End Sub

Private Sub Command3_Click()
    Form5.Visible = True
    Form4.Visible = False
End Sub

Private Sub Command4_Click()
    Form2.Visible = True
    Form4.Visible = False
End Sub
