VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   0  'None
   Caption         =   "Menu3"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form28"
   Picture         =   "Form13.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voltar"
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
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Novo"
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
      TabIndex        =   1
      Top             =   2520
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editar"
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
      TabIndex        =   0
      Top             =   3360
      Width           =   1170
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HC0&
    Label2.ForeColor = &HC0&
    Label3.ForeColor = &HC0&
End Sub

Private Sub Label1_Click()
    Load Form10
    Load Form11
    Load Form12
    reg.nome = ""
    reg.raca = ""
    reg.sexo = ""
    reg.foto = ""
    reg.ouro = 0
    reg.vida = 0
    reg.mana = 0
    reg.pvida = 0
    reg.pmana = 0
    reg.pc = 0
    reg.exp = 0
    reg.ar = False
    reg.terra = False
    reg.fogo = False
    reg.agua = False
    reg.forca = False
    reg.natureza = False
    reg.meta = False
    reg.nnegra = False
    reg.nbranca = False
    reg.invoca = False
    Form10.Visible = True
    Form13.Visible = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF&
End Sub

Private Sub Label2_Click()
    pathfx1 = ""
    pathfx1 = InputBox("Insira o nome da personagem", "Editar")
    If pathfx1 <> "" Then
        If Dir(path & "Personagens\" & pathfx1 & ".dat") <> "" Then
            pathfx1 = path & "Personagens\" & pathfx1 & ".dat"
            Open pathfx1 For Random As #1 Len = Len(reg)
            Get #1, 1, reg
            Form14.Text1.Text = RTrim(reg.nome)
            Form14.Image1.Picture = LoadPicture(reg.foto)
            Select Case RTrim(reg.raca)
                Case "h"
                    Form14.Text8.Text = "Humano"
                Case "d"
                    Form14.Text8.Text = "Anao"
                Case "e"
                    Form14.Text8.Text = "Elfo"
                Case "he"
                    Form14.Text8.Text = "Meio-Elfo"
                Case "g"
                    Form14.Text8.Text = "Gnomo"
                Case "ha"
                    Form14.Text8.Text = "Halfling"
                Case "ho"
                    Form14.Text8.Text = "Meio-Orco"
                Case "hog"
                    Form14.Text8.Text = "Meio-Ogre"
            End Select
            If reg.sexo = "m" Then
                Form14.Text9.Text = "Masculino"
            Else
                Form14.Text9.Text = "Feminino"
            End If
            Form14.Text10.Text = reg.pvida
            Form14.Text11.Text = reg.pmana
            Form14.Text2.Text = reg.ouro
            Form14.Text3.Text = reg.vida
            Form14.Text4.Text = reg.mana
            Form14.Text5.Text = reg.exp
            Form14.Text6.Text = reg.pc
            Form14.Text7.Text = Int(reg.exp / 1000) + 1
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
            Close #1
            Form14.Visible = True
            Form13.Visible = False
        Else
            MsgBox "Não encontrado", vbOKOnly, "Editar"
        End If
    End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &HFF&
End Sub

Private Sub Label3_Click()
    Form9.Visible = True
    Form13.Visible = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.ForeColor = &HFF&
End Sub
