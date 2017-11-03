VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Compras"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form27"
   MousePointer    =   99  'Custom
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
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
      Left            =   7800
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         Top             =   1560
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
         TabIndex        =   19
         Text            =   "0"
         Top             =   720
         Width           =   975
      End
      Begin VB.Image Image18 
         Height          =   480
         Left            =   2280
         Picture         =   "Form12.frx":0000
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image17 
         Height          =   480
         Left            =   120
         Picture         =   "Form12.frx":0442
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   2280
         Picture         =   "Form12.frx":0884
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image15 
         Height          =   480
         Left            =   120
         Picture         =   "Form12.frx":0CC6
         Top             =   600
         Width           =   480
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
         TabIndex        =   18
         Top             =   1200
         Width           =   555
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
         TabIndex        =   17
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "500"
      Top             =   2250
      Visible         =   0   'False
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "10"
      Top             =   6480
      Width           =   975
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "10"
      Top             =   4410
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "10"
      Top             =   3570
      Width           =   1695
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "500"
      Top             =   2730
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
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
      Height          =   1695
      Left            =   7800
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox Text8 
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
         TabIndex        =   22
         Text            =   "0"
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   2280
         Picture         =   "Form12.frx":1108
         Top             =   600
         Width           =   480
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   120
         Picture         =   "Form12.frx":154A
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.Image Image11 
      Height          =   705
      Left            =   9120
      Picture         =   "Form12.frx":198C
      Stretch         =   -1  'True
      Top             =   8040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label12 
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
      Left            =   7680
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Image Image10 
      Height          =   750
      Left            =   6720
      Picture         =   "Form12.frx":6E42
      Top             =   8040
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image9 
      Height          =   750
      Left            =   5040
      Picture         =   "Form12.frx":8101
      Top             =   8040
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   360
      Left            =   3000
      Picture         =   "Form12.frx":93A7
      Top             =   4200
      Width           =   360
   End
   Begin VB.Image Image7 
      Height          =   360
      Left            =   3000
      Picture         =   "Form12.frx":949B
      Top             =   4560
      Width           =   360
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   3000
      Picture         =   "Form12.frx":984E
      Top             =   3360
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   360
      Left            =   3000
      Picture         =   "Form12.frx":9942
      Top             =   3720
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   3000
      Picture         =   "Form12.frx":9CF5
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   3000
      Picture         =   "Form12.frx":A0A8
      Top             =   2520
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   5460
      Left            =   7440
      Picture         =   "Form12.frx":A19C
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3840
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   6240
      Picture         =   "Form12.frx":47A32
      Top             =   1200
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos Restantes"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   6480
      Width           =   2535
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
      Left            =   3720
      TabIndex        =   11
      Top             =   2760
      Width           =   2385
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
      Left            =   3720
      TabIndex        =   10
      Top             =   3600
      Width           =   2340
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
      Left            =   3720
      TabIndex        =   9
      Top             =   4440
      Width           =   2505
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
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   675
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
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   720
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
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
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
      Left            =   8760
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos de Compra"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ouro, vida, mana, poções..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   4740
   End
   Begin VB.Image Image12 
      Height          =   6060
      Left            =   960
      Picture         =   "Form12.frx":48C87
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   4560
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image10_Click()
    Label3.Visible = False
    Label2.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Text3.Visible = True
    Text4.Visible = True
    Image2.Visible = True
    Image3.Visible = True
    Image4.Visible = True
    Image5.Visible = True
    Image6.Visible = True
    Image7.Visible = True
    Image8.Visible = True
    Image9.Visible = True
    Image12.Visible = False
    Text5.Visible = False
    Frame1.Visible = False
    Frame2.Visible = False
    Image10.Visible = False
    Image11.Visible = False
    Label12.Visible = False
    Text1.Text = Text5.Text
    Text1.TabIndex = 1
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.BorderStyle = 1
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.BorderStyle = 0
End Sub

Sub gravar()
    pathfx1 = path & "Personagens\" & pathfx1 & ".dat"
    Open pathfx1 For Output As #1 Len = Len(reg)
    reg.nome = RTrim(reg.nome)
    reg.raca = RTrim(reg.raca)
    reg.sexo = RTrim(reg.sexo)
    reg.foto = RTrim(reg.foto)
    reg.ouro = Val(Text5.Text)
    reg.vida = Val(Text2.Text)
    reg.mana = Val(Text3.Text)
    reg.pvida = Val(Text6.Text)
    reg.pmana = Val(Text7.Text)
    reg.pc = Val(Text4.Text)
    reg.exp = Val(Text8.Text)
    Close #1
    Open pathfx1 For Random As #1 Len = Len(reg)
    Put #1, 1, reg
    Close #1
    MsgBox "Ficheiro gravado como " & pathfx1, vbOKOnly, "Sucesso"
    Form13.Visible = True
    Unload Form10
    Unload Form11
    Unload Form12
End Sub
Private Sub Image11_Click()
    pathfx1 = InputBox("Insira o nome do ficheiro", "Gravar Personagem", RTrim(reg.nome))
    If pathfx1 <> "" Then
        If Dir(path & "Personagens\" & pathfx1 & ".dat") <> "" Then
            bt = MsgBox("Ficheiro existente. Deseja substituir?", vbYesNo, "AVISO!")
            If bt = 6 Then
                gravar
            End If
        Else
            gravar
        End If
    End If
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.BorderStyle = 1
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.BorderStyle = 0
End Sub

Private Sub Image13_Click()
    If Val(Text8.Text) >= 1 Then
        Text8.Text = Val(Text8.Text) - 1
        Text5.Text = Val(Text5.Text) + 1
        gold = gold + 1
    End If
End Sub

Private Sub Image14_Click()
    If Val(Text5.Text) >= 1 Then
        Text8.Text = Val(Text8.Text) + 1
        Text5.Text = Val(Text5.Text) - 1
        gold = gold - 1
    End If
End Sub

Private Sub Image15_Click()
    If Val(Text6.Text) >= 1 Then
        Text6.Text = Val(Text6.Text) - 1
        Text5.Text = Val(Text5.Text) + 20
        gold = gold + 20
    End If
End Sub

Private Sub Image16_Click()
    If Val(Text5.Text) >= 20 Then
        Text6.Text = Val(Text6.Text) + 1
        Text5.Text = Val(Text5.Text) - 20
        gold = gold - 20
    End If
End Sub

Private Sub Image17_Click()
    If Val(Text7.Text) >= 1 Then
        Text7.Text = Val(Text7.Text) - 1
        Text5.Text = Val(Text5.Text) + 20
        gold = gold + 20
    End If
End Sub

Private Sub Image18_Click()
    If Val(Text5.Text) >= 20 Then
        Text7.Text = Val(Text7.Text) + 1
        Text5.Text = Val(Text5.Text) - 20
        gold = gold - 20
    End If
End Sub

Private Sub Image3_Click()
    If Val(Text4.Text) >= 1 Then
        Text1.Text = Val(Text1.Text) + 5
        Text4.Text = Val(Text4.Text) - 1
    End If
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.BorderStyle = 1
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.BorderStyle = 0
End Sub

Private Sub Image4_Click()
    If Val(Text1.Text) > gold Then
        Text1.Text = Val(Text1.Text) - 5
        Text4.Text = Val(Text4.Text) + 1
    End If
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.BorderStyle = 1
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.BorderStyle = 0
End Sub
Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.BorderStyle = 1
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.BorderStyle = 0
End Sub

Private Sub Image6_Click()
    If Val(Text4.Text) > 0 Then
        Text2.Text = Val(Text2.Text) + 1
        Text4.Text = Val(Text4.Text) - 1
    End If
End Sub

Private Sub Image5_Click()
    If Val(Text2.Text) > 10 Then
        Text2.Text = Val(Text2.Text) - 1
        Text4.Text = Val(Text4.Text) + 1
    End If
End Sub

Private Sub Image8_Click()
    If Val(Text4.Text) > 0 Then
        Text3.Text = Val(Text3.Text) + 1
        Text4.Text = Val(Text4.Text) - 1
    End If
End Sub

Private Sub Image7_Click()
    If Val(Text3.Text) > 10 Then
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

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.BorderStyle = 1
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.BorderStyle = 0
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.BorderStyle = 1
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image8.BorderStyle = 0
End Sub

Private Sub Image9_Click()
    Label3.Visible = True
    Label2.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    Image4.Visible = False
    Image5.Visible = False
    Image6.Visible = False
    Image7.Visible = False
    Image8.Visible = False
    Image9.Visible = False
    Image12.Visible = True
    Text5.Visible = True
    Frame1.Visible = True
    Frame2.Visible = True
    Image10.Visible = True
    Image11.Visible = True
    Label12.Visible = True
    Text5.Text = Text1.Text
    Text5.TabIndex = 1
End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.BorderStyle = 1
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.BorderStyle = 0
End Sub

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.BorderStyle = 1
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.BorderStyle = 0
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.BorderStyle = 1
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.BorderStyle = 0
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.BorderStyle = 1
End Sub

Private Sub Image15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.BorderStyle = 0
End Sub

Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.BorderStyle = 1
End Sub

Private Sub Image16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.BorderStyle = 0
End Sub

Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image17.BorderStyle = 1
End Sub

Private Sub Image17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image17.BorderStyle = 0
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image18.BorderStyle = 1
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image18.BorderStyle = 0
End Sub

