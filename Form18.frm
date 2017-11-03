VERSION 5.00
Begin VB.Form Form18 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Editar2"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form18"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      ItemData        =   "Form18.frx":0000
      Left            =   8400
      List            =   "Form18.frx":0002
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form18.frx":0004
      Left            =   8400
      List            =   "Form18.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Height          =   1095
      Left            =   8640
      Picture         =   "Form18.frx":0008
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Pontos de Experiência "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   4200
      TabIndex        =   10
      Top             =   5520
      Width           =   2895
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   120
         Picture         =   "Form18.frx":513A
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   2280
         Picture         =   "Form18.frx":557C
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2490
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3330
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4170
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Caption         =   "Poções "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   840
      TabIndex        =   1
      Top             =   5160
      Width           =   2895
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1125
         TabIndex        =   5
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1110
         TabIndex        =   4
         Top             =   1200
         Width           =   555
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   120
         Picture         =   "Form18.frx":59BE
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   2280
         Picture         =   "Form18.frx":5E00
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   120
         Picture         =   "Form18.frx":6242
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   2280
         Picture         =   "Form18.frx":6684
         Top             =   1440
         Width           =   480
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Escolas de Magia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Editar Personagem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   19
      Top             =   360
      Width           =   3270
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ouro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   18
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mana"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   17
      Top             =   4200
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Vida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   16
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1 Ponto = 1 Mana"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   15
      Top             =   4200
      Width           =   2505
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1 Ponto = 1 Vida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   14
      Top             =   3360
      Width           =   2340
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1 Ponto = 5 Ouro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   13
      Top             =   2520
      Width           =   2385
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos de Compra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   12
      Top             =   1680
      Width           =   2595
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3480
      Picture         =   "Form18.frx":6AC6
      Top             =   2280
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   3480
      Picture         =   "Form18.frx":6BBA
      Top             =   2640
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   3480
      Picture         =   "Form18.frx":6F6D
      Top             =   3480
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   3480
      Picture         =   "Form18.frx":7320
      Top             =   3120
      Width           =   360
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   3480
      Picture         =   "Form18.frx":7414
      Top             =   4320
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   360
      Left            =   3480
      Picture         =   "Form18.frx":77C7
      Top             =   3960
      Width           =   360
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    List1.Clear
    Select Case Combo1.Text
        Case "Ar"
            List1.AddItem "Vitalidade do Ar"
            List1.AddItem "Vapores Venenosos"
            List1.AddItem "Chamar Ventos"
            List1.AddItem "Corpo de Ar"
            List1.AddItem "Elementar do Ar"
            temp = Combo1.Text
        Case "Terra"
            List1.AddItem "Força da Terra"
            List1.AddItem "Lançar Pedra"
            List1.AddItem "Barreira de Pedra"
            List1.AddItem "Corpo de Pedra"
            List1.AddItem "Elementar do Pedra"
            temp = Combo1.Text
        Case "Fogo"
            List1.AddItem "Agilidade do Fogo"
            List1.AddItem "Barreira de Fogo"
            List1.AddItem "Bola de Fogo"
            List1.AddItem "Corpo de Fogo"
            List1.AddItem "Elementar do Fogo"
            temp = Combo1.Text
        Case "Água"
            List1.AddItem "Pureza de Água"
            List1.AddItem "Chamar Névoa"
            List1.AddItem "Tempestade de Gelo"
            List1.AddItem "Corpo de Água"
            List1.AddItem "Elementar do Água"
            temp = Combo1.Text
        Case "Força"
            List1.AddItem "Escudo Protector"
            List1.AddItem "Raio Eléctrico"
            List1.AddItem "Barreira de Força"
            List1.AddItem "Trovão"
            List1.AddItem "Desintegrar"
            temp = Combo1.Text
        Case "Natureza"
            List1.AddItem "Encantar Criatura"
            List1.AddItem "Floresta"
            List1.AddItem "Controlar Criatura"
            List1.AddItem "Invocar Criatura"
            List1.AddItem "Regeneração"
            temp = Combo1.Text
        Case "Metamorfose"
            List1.AddItem "Resistir Magia"
            List1.AddItem "Dispersar Magias"
            List1.AddItem "Enfraquecer"
            List1.AddItem "Escudo Reflector"
            List1.AddItem "Polymorph"
            temp = Combo1.Text
        Case "Necro. Negra"
            List1.AddItem "Prejudicar"
            List1.AddItem "Conjurar Espírito"
            List1.AddItem "Invocar Zombie"
            List1.AddItem "Criar Zombie"
            List1.AddItem "Sugar Vida"
            temp = Combo1.Text
        Case "Necro. Branca"
            List1.AddItem "Curativo"
            List1.AddItem "Desentoxicar"
            List1.AddItem "Curar"
            List1.AddItem "Santuário"
            List1.AddItem "Ressuscitar"
            temp = Combo1.Text
        Case "Invocação"
            List1.AddItem "Troll"
            List1.AddItem "Orco"
            List1.AddItem "Ogre"
            List1.AddItem "Almas Perdidas"
            List1.AddItem "Portas do Inferno"
            temp = Combo1.Text
    End Select
End Sub

Private Sub Command1_Click()
    reg.pc = Val(Text4.Text)
    reg.ouro = Val(Text1.Text)
    reg.vida = Val(Text2.Text)
    reg.mana = Val(Text3.Text)
    reg.pvida = Val(Text5.Text)
    reg.pmana = Val(Text6.Text)
    reg.exp = Val(Text7.Text)
    Select Case temp
        Case "Ar"
            reg.ar = True
        Case "Terra"
            reg.terra = True
        Case "Fogo"
            reg.fogo = True
        Case "Água"
            reg.agua = True
        Case "Força"
            reg.forca = True
        Case "Natureza"
            reg.natureza = True
        Case "Metamorfose"
            reg.meta = True
        Case "Necro. Negra"
            reg.nnegra = True
        Case "Necro. Branca"
            reg.nbranca = True
        Case "Invocação"
            reg.invoca = True
    End Select
    Open pathfx1 For Random As #1 Len = Len(reg)
    Put #1, 1, reg
    Close #1
    Form14.Text10.Text = reg.pvida
    Form14.Text11.Text = reg.pmana
    Form14.Text2.Text = reg.ouro
    Form14.Text3.Text = reg.vida
    Form14.Text4.Text = reg.mana
    Form14.Text5.Text = reg.exp
    Form14.Text6.Text = reg.pc
    Form14.List1.Clear
    If reg.ar = True Then
        Form14.List1.AddItem "Ar"
    End If
    If reg.terra = True Then
        Form14.List1.AddItem "Terra"
    End If
    If reg.fogo = True Then
        Form14.List1.AddItem "Fogo"
    End If
    If reg.agua = True Then
        Form14.List1.AddItem "Água"
    End If
    If reg.forca = True Then
        Form14.List1.AddItem "Força"
    End If
    If reg.natureza = True Then
        Form14.List1.AddItem "Natureza"
    End If
    If reg.meta = True Then
        Form14.List1.AddItem "Metamorfose"
    End If
    If reg.nnegra = True Then
        Form14.List1.AddItem "Necro. Negra"
    End If
    If reg.nbranca = True Then
        Form14.List1.AddItem "Necro. Branca"
    End If
    If reg.invoca = True Then
        Form14.List1.AddItem "Invocação"
    End If
    List1.Visible = False
    Combo1.Visible = False
    Label2.Visible = False
    Form14.Visible = True
    Form18.Visible = False
End Sub

Private Sub Image1_Click()
    If Val(Text4.Text) >= 1 Then
        Text4.Text = Val(Text4.Text) - 1
        Text1.Text = Val(Text1.Text) + 5
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.BorderStyle = 1
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.BorderStyle = 0
End Sub

Private Sub Image10_Click()
    If Val(Text6.Text) > reg.pmana Then
        Text6.Text = Val(Text6.Text) - 1
        Text1.Text = Val(Text1.Text) + 20
        gold = gold + 20
    End If
End Sub

Private Sub Image11_Click()
    If Val(Text1.Text) > 0 Then
        Text7.Text = Val(Text7.Text) + 1
        Text1.Text = Val(Text1.Text) - 1
        gold = gold - 1
    End If
End Sub

Private Sub Image12_Click()
    If Val(Text7.Text) > reg.exp Then
        Text7.Text = Val(Text7.Text) - 1
        Text1.Text = Val(Text1.Text) + 1
        gold = gold + 1
    End If
End Sub

Private Sub Image2_Click()
    If Val(Text1.Text) > gold Then
        Text1.Text = Val(Text1.Text) - 5
        Text4.Text = Val(Text4.Text) + 1
    End If
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.BorderStyle = 0
End Sub

Private Sub Image3_Click()
    If Val(Text4.Text) >= 1 Then
        Text4.Text = Val(Text4.Text) - 1
        Text2.Text = Val(Text2.Text) + 1
    End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.BorderStyle = 0
End Sub

Private Sub Image4_Click()
    If Val(Text2.Text) > reg.vida Then
        Text2.Text = Val(Text2.Text) - 1
        Text4.Text = Val(Text4.Text) + 1
    End If
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.BorderStyle = 1
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.BorderStyle = 0
End Sub

Private Sub Image5_Click()
    If Val(Text4.Text) >= 1 Then
        Text4.Text = Val(Text4.Text) - 1
        Text3.Text = Val(Text3.Text) + 1
    End If
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.BorderStyle = 1
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.BorderStyle = 0
End Sub

Private Sub Image6_Click()
    If Val(Text3.Text) > reg.mana Then
        Text3.Text = Val(Text3.Text) - 1
        Text4.Text = Val(Text4.Text) + 1
    End If
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.BorderStyle = 1
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.BorderStyle = 0
End Sub

Private Sub Image7_Click()
    If Val(Text1.Text) >= 20 Then
        Text1.Text = Val(Text1.Text) - 20
        Text5.Text = Val(Text5.Text) + 1
        gold = gold - 20
    End If
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.BorderStyle = 1
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.BorderStyle = 0
End Sub

Private Sub Image8_Click()
    If Val(Text5.Text) > reg.pvida Then
        Text5.Text = Val(Text5.Text) - 1
        Text1.Text = Val(Text1.Text) + 20
        gold = gold + 20
    End If
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.BorderStyle = 1
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.BorderStyle = 0
End Sub

Private Sub Image9_Click()
    If Val(Text1.Text) >= 20 Then
        Text1.Text = Val(Text1.Text) - 20
        Text6.Text = Val(Text6.Text) + 1
        gold = gold - 20
    End If
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.BorderStyle = 1
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.BorderStyle = 0
End Sub
Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.BorderStyle = 1
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.BorderStyle = 0
End Sub
Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.BorderStyle = 1
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.BorderStyle = 0
End Sub
Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image12.BorderStyle = 1
End Sub

Private Sub Image12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image12.BorderStyle = 0
End Sub

