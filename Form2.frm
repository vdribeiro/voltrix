VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Menu1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   Picture         =   "Form2.frx":030A
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Curiosidades"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   4920
      TabIndex        =   5
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Créditos"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5400
      TabIndex        =   4
      Top             =   4680
      Width           =   1635
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instruções"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5160
      TabIndex        =   3
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jogar"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5640
      TabIndex        =   2
      Top             =   1560
      Width           =   1065
   End
   Begin VB.OLE OLE1 
      Class           =   "MPlayer"
      Height          =   30
      Left            =   120
      OleObjectBlob   =   "Form2.frx":15FC4C
      SizeMode        =   1  'Stretch
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   5760
      TabIndex        =   0
      Top             =   5640
      Width           =   765
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
    Unload Form1
    End
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF&
End Sub

Private Sub Label2_Click()
    Form9.Visible = True
    Form2.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &HFF&
End Sub

Private Sub Label3_Click()
    Form8.Visible = True
    Form8.Command1.Visible = False
    Form8.Command2.Visible = False
    Form8.Label1(0).Alignment = 2
    Form8.Label1(1).Alignment = 2
    Form8.Label1(2).Alignment = 2
    Form8.Label1(3).Alignment = 2
    Form8.Label1(4).Alignment = 2
    Form8.Label1(5).Alignment = 2
    Form8.Label1(6).Alignment = 2
    Form8.Label1(7).Alignment = 2
    Form8.Label1(8).Alignment = 2
    Form8.Label1(9).Alignment = 2
    Form8.Label2.Caption = "Instruções"
    Form8.Label1(0).Caption = "- ESCOLHA UMA OPÇÃO DO MENU -"
    Form8.Label1(1).Caption = ""
    Form8.Label1(2).Caption = ""
    Form8.Label1(3).Caption = ""
    Form8.Label1(4).Caption = ""
    Form8.Label1(5).Caption = ""
    Form8.Label1(6).Caption = ""
    Form8.Label1(7).Caption = ""
    Form8.Label1(8).Caption = ""
    Form8.Label1(9).Caption = ""
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &HFF&
End Sub

Private Sub Label4_Click()
    Form3.Visible = True
    Form2.Visible = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label4.ForeColor = &HFF&
End Sub

Private Sub Label5_Click()
    Form4.Visible = True
    Form2.Visible = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.ForeColor = &HFF&
End Sub

Private Sub Form_Load()

    '-------------------------------------------------------------------------------------------------
    'DEFINIÇÃO DA PATH
    path = "C:\Voltrix\"
    '-------------------------------------------------------------------------------------------------
    
    '-------------------------------------------------------------------------------------------------
    'MUSICA DO MENU
    'OLE1.DoVerb
    '-------------------------------------------------------------------------------------------------
    
    If Dir(path & "Personagens\cpu.dat") = "" Then
        Open path & "Personagens\cpu.dat" For Output As #1 Len = Len(reg)
        Close #1
        Open path & "Personagens\cpu.dat" For Random As #1 Len = Len(reg)
        reg.nome = "CPU"
        reg.raca = "h"
        reg.sexo = "m"
        reg.foto = path & "Faces\cpu.jpg"
        reg.ouro = 500
        reg.vida = 10
        reg.mana = 10
        reg.pvida = 0
        reg.pmana = 0
        reg.pc = 10
        reg.exp = 0
        Put #1, 1, reg
        Close #1
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HC0&
    Label2.ForeColor = &HC0&
    Label3.ForeColor = &HC0&
    Label4.ForeColor = &HC0&
    Label5.ForeColor = &HC0&
End Sub

