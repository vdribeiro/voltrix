VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Instru��es"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   10440
      Picture         =   "Form8.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   240
      Picture         =   "Form8.frx":12B6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   10
      Top             =   5160
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   8
      Top             =   3720
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "- ESCOLHA UMA OP��O DO MENU -"
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   11895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   1
      Top             =   4440
      Width           =   11895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INSTRU��ES"
      BeginProperty Font 
         Name            =   "Digiface"
         Size            =   44.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1080
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   23
      Left            =   10800
      Picture         =   "Form8.frx":255C
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   22
      Left            =   8640
      Picture         =   "Form8.frx":3C3B
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   21
      Left            =   6480
      Picture         =   "Form8.frx":531A
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   20
      Left            =   4320
      Picture         =   "Form8.frx":69F9
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   19
      Left            =   2160
      Picture         =   "Form8.frx":80D8
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   18
      Left            =   0
      Picture         =   "Form8.frx":97B7
      Top             =   6480
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   17
      Left            =   10800
      Picture         =   "Form8.frx":AE96
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   16
      Left            =   8640
      Picture         =   "Form8.frx":C575
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   15
      Left            =   6480
      Picture         =   "Form8.frx":DC54
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   14
      Left            =   4320
      Picture         =   "Form8.frx":F333
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   13
      Left            =   2160
      Picture         =   "Form8.frx":10A12
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   12
      Left            =   0
      Picture         =   "Form8.frx":120F1
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   11
      Left            =   10800
      Picture         =   "Form8.frx":137D0
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   10
      Left            =   8640
      Picture         =   "Form8.frx":14EAF
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   9
      Left            =   6480
      Picture         =   "Form8.frx":1658E
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   8
      Left            =   4320
      Picture         =   "Form8.frx":17C6D
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   7
      Left            =   2160
      Picture         =   "Form8.frx":1934C
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   6
      Left            =   0
      Picture         =   "Form8.frx":1AA2B
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   5
      Left            =   10800
      Picture         =   "Form8.frx":1C10A
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   4
      Left            =   8640
      Picture         =   "Form8.frx":1D7E9
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   3
      Left            =   6480
      Picture         =   "Form8.frx":1EEC8
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   2
      Left            =   4320
      Picture         =   "Form8.frx":205A7
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   1
      Left            =   2160
      Picture         =   "Form8.frx":21C86
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   2160
      Index           =   0
      Left            =   0
      Picture         =   "Form8.frx":23365
      Top             =   0
      Width           =   2160
   End
   Begin VB.Menu a 
      Caption         =   "Voltrix"
      Begin VB.Menu b 
         Caption         =   "A Lenda"
         Index           =   0
      End
      Begin VB.Menu b 
         Caption         =   "Objectivos"
         Index           =   1
      End
   End
   Begin VB.Menu c 
      Caption         =   "Personagens"
      Begin VB.Menu d 
         Caption         =   "Cria��o"
         Index           =   0
      End
      Begin VB.Menu d 
         Caption         =   "Edita��o"
         Index           =   1
      End
   End
   Begin VB.Menu e 
      Caption         =   "Combates"
      Begin VB.Menu f 
         Caption         =   "Prepara��o"
         Index           =   0
      End
      Begin VB.Menu f 
         Caption         =   "O Duelo"
         Index           =   1
      End
   End
   Begin VB.Menu g 
      Caption         =   "Gloss�rio"
   End
   Begin VB.Menu h 
      Caption         =   "Fechar"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub align()
    Label1(0).Alignment = 2
    Label1(1).Alignment = 2
    Label1(2).Alignment = 2
    Label1(3).Alignment = 2
    Label1(4).Alignment = 2
    Label1(5).Alignment = 2
    Label1(6).Alignment = 2
    Label1(7).Alignment = 2
    Label1(8).Alignment = 2
    Label1(9).Alignment = 2
End Sub

Sub duel()
    Command1.Visible = False
    Command2.Visible = True
    Label2.Caption = "O Duelo"
    Label1(0).Caption = "A parte mas importante e em que consiste o jogo s�o os duelos. Nos combates, os"
    Label1(1).Caption = "combatentes nunca se tocam; usam as suas criaturas e os seus feiti�os para lutarem."
    Label1(2).Caption = "O combate acaba quando um jogador for derrotado. �, de seguida, apresentado um quadro"
    Label1(3).Caption = "com detalhes do que foi ganho/perdido. O combate � feito por turnos - ataque e defesa."
    Label1(4).Caption = "Os turnos s�o alternados da seguinte maneira: o teu turno de ataque, o turno de defesa do"
    Label1(5).Caption = "inimigo, o turno de ataque do inimigo, o teu turno de defesa. O 1� ataque pertence ao 1�"
    Label1(6).Caption = "jogador. Nos turnos de ataque, o jogador pode lan�ar feiti�os e/ou invocar criaturas (4 no"
    Label1(7).Caption = "m�ximo). Nos turnos de defesa, o jogador s� pode controlar as suas criaturas em jogo e"
    Label1(8).Caption = "lan�ar feiti�os defensivos. Todos os comandos do turno presente s� s�o executados"
    Label1(9).Caption = "quando esse turno acabar. Em rela��o �s criaturas: ir� aparecer ao cimo da imagem da"
End Sub

Sub legend()
    Command1.Visible = False
    Command2.Visible = True
    Label2.Caption = "A Lenda"
    Label1(0).Caption = "No ano de 1885, com o progressivo aparecimento de novas tecnologias, o uso de magias"
    Label1(1).Caption = "entrou em decad�ncia. As pessoas j� acreditavam no progresso e foram esquecendo as"
    Label1(2).Caption = "maravilhas m�gicas. A tecnologia come�ou a ganhar novo interesse e com o aparecimento"
    Label1(3).Caption = "do primeiro comboio a magia come�ou a ter um efeito t�o negativo que foi proibido o uso."
    Label1(4).Caption = "Tal decis�o causou uma guerra sangrenta em que a tecnologia triunfou gra�as �s armas"
    Label1(5).Caption = "de fogo - as novas maravilhas criadas pelo homem. No entanto, esta vit�ria s� seria"
    Label1(6).Caption = "completa com a destrui��o do templo Volknor - o lugar sagrado onde todas as magias"
    Label1(7).Caption = "culminavam. S� os grandes magos conseguiam entrar no templo. A destrui��o estava"
    Label1(8).Caption = "preparada e tudo pronto para rebentar. Numa frac��o de segundo antes da explos�o, tudo"
    Label1(9).Caption = "parou, o templo desapareceu e a mem�ria das pessoas ficou em branco. Ninguem se"
End Sub

Private Sub b_Click(Index As Integer)
    align
    If Index = 0 Then
        legend
    Else
        Command1.Visible = False
        Command2.Visible = False
        Label2.Caption = "Objectivos"
        Label1(0).Caption = "O objectivo do jogo � chegar ao n�vel 20 dando depois ao veterano a possibilidade de ir"
        Label1(1).Caption = "buscar uma pedra Voltrix que lhe dar� acesso a tudo do jogo, ou seja, todas as magias,"
        Label1(2).Caption = "pontos sem limite, etc� Torna-se um verdadeiro deus!"
        Label1(3).Caption = ""
        Label1(4).Caption = ""
        Label1(5).Caption = ""
        Label1(6).Caption = ""
        Label1(7).Caption = ""
        Label1(8).Caption = ""
        Label1(9).Caption = ""
    End If
End Sub

Private Sub Command1_Click()
    Command1.Visible = False
    Command2.Visible = True
    Select Case Label2.Caption
        Case "A Lenda"
            legend
        Case "O Duelo"
            duel
    End Select
End Sub

Private Sub Command2_Click()
    Command2.Visible = False
    Command1.Visible = True
    Select Case Label2.Caption
        Case "A Lenda"
            Label1(0).Caption = "lembrava do que acontecera. S�culo e meio depois deste acontecimento, o templo"
            Label1(1).Caption = "reaparece e traz consigo as for�as do passado. Nesta altura, o n� de praticantes de"
            Label1(2).Caption = "magia � bastante baixo. Contudo, a sua hist�ria continuara viva gra�as a um mago que"
            Label1(3).Caption = "escrevera sobre o templo. Os �ltimos praticantes para l� correram na esperan�a de"
            Label1(4).Caption = "entrarem: �S� os feiticeiros mais poderosos poder�o entrar!� - ressoava uma voz. Foi"
            Label1(5).Caption = "ent�o realizado um torneio onde se decidiria quem o iria fazer. Os vencedores poderiam"
            Label1(6).Caption = "entrar no templo e transformarem-se em deuses, dominando todas as artes m�gicas."
            Label1(7).Caption = "Quem vencer�? Vemo-nos na arena!!!"
            Label1(8).Caption = ""
            Label1(9).Caption = ""
        Case "O Duelo"
            Label1(0).Caption = "criatura correspondente o poder de ataque/vida. As criaturas poder�o ser comandadas"
            Label1(1).Caption = "fazendo duplo clique na criatura desejada. Quando a vida desta �ltima chegar a zero,"
            Label1(2).Caption = "vai para o cemit�rio onde poder� ser ressuscitada. Quando tiveres 4 criaturas em jogo,"
            Label1(3).Caption = "se desejares invocar outra criatura, a 1� ser� sacrificada dando lugar � nova. Existem"
            Label1(4).Caption = "magias que mexem com as caracter�sticas das criaturas, no entanto, a criatura nunca pode"
            Label1(5).Caption = "ter mais do que 9/9 ou menos que 1/1 com a exep��o do �Polymorph� em que a criatura"
            Label1(6).Caption = "fica com 0/0 (� transformada numa ovelha). Com a pr�tica, aprender�s o resto. Boa sorte!"
            Label1(7).Caption = ""
            Label1(8).Caption = ""
            Label1(9).Caption = ""
    End Select
End Sub

Private Sub d_Click(Index As Integer)
    align
    If Index = 0 Then
        Command1.Visible = False
        Command2.Visible = False
        Label2.Caption = "Cria��o"
        Label1(0).Caption = "Esta op��o � indispens�vel antes de come�ar o combate. Cada jogador tem de criar uma"
        Label1(1).Caption = "personagem, para poder jogar, e desenvolve-la. Ganhando combates, o jogador adquire"
        Label1(2).Caption = "ouro e pontos de experi�ncia. Com o ouro poder� comprar po��es de vida e mana ou at�"
        Label1(3).Caption = "fazer um pouco de batota e comprar experi�ncia. Os pontos de experi�ncia permitem o"
        Label1(4).Caption = "jogador avan�ar de n�vel ganhando assim 2 pontos de compra os quais o possibilitar�o"
        Label1(5).Caption = "de melhorar o status da personagem. 1 ponto de compra equivale ou a +5 de ouro ou +1"
        Label1(6).Caption = "ponto de vida ou +1 ponto de mana. No decurso da constru��o da personagem, ir�s tamb�m"
        Label1(7).Caption = "defrontar-te com a escolha de um col�gio de magia. Ter�s, inicialmente, de optar por um"
        Label1(8).Caption = "dos 10 que te ser�o apresentados. Cada col�gio cont�m 5 magias. Acabada a cria��o, j�"
        Label1(9).Caption = "podes combater, ou se te enganaste nalguma coisa ainda tens a op��o �EDITAR�."
    Else
        Command1.Visible = False
        Command2.Visible = False
        Label2.Caption = "Edita��o"
        Label1(0).Caption = "Depois de criada uma personagem, podes edit�-la. Quanto tiveres pontos de compra para"
        Label1(1).Caption = "gastar, ou quiseres comprar mais po��es ou pontos de experi�ncia, ou at� quando te for"
        Label1(2).Caption = "dada a possibilidade de obteres mais um col�gio de magia, ter�s de ir � op��o de edi��o."
        Label1(3).Caption = "Modificar a personagem � vital para a sua sobreviv�ncia. Sempre que possa, edite-a."
        Label1(4).Caption = ""
        Label1(5).Caption = ""
        Label1(6).Caption = ""
        Label1(7).Caption = ""
        Label1(8).Caption = ""
        Label1(9).Caption = ""
    End If
End Sub

Private Sub f_Click(Index As Integer)
    align
    If Index = 0 Then
        Command1.Visible = False
        Command2.Visible = False
        Label2.Caption = "Prepara��o"
        Label1(0).Caption = "Tenha calma, relaxe. Aproveite os extras e escolha uma arena, o m�todo de distribui��o de"
        Label1(1).Caption = "cartas e uma musiquinha de fundo ao seu gosto. A arena e a m�sica n�o influenciam a"
        Label1(2).Caption = "batalha, mas o m�todo de como as cartas s�o distribuidas - isso sim. Existem dois"
        Label1(3).Caption = "m�todos: O Tradicional - no inicio do turno � acrescentada uma carta � m�o do jogador, no"
        Label1(4).Caption = "caso de ter a m�o cheia (7 cartas), � descartada a ultima carta. Existe tamb�m o m�todo"
        Label1(5).Caption = "aleat�rio que consite em renovar toda a m�o do jogador no inicio do turno. Quando estiver"
        Label1(6).Caption = "pronto - ...n�o tenha piedade!"
        Label1(7).Caption = ""
        Label1(8).Caption = ""
        Label1(9).Caption = ""
    Else
        duel
    End If
End Sub

Private Sub g_Click()
    Label1(0).Alignment = 0
    Label1(1).Alignment = 0
    Label1(2).Alignment = 0
    Label1(3).Alignment = 0
    Label1(4).Alignment = 0
    Label1(5).Alignment = 0
    Label1(6).Alignment = 0
    Label1(7).Alignment = 0
    Label1(8).Alignment = 0
    Label1(9).Alignment = 0
    Command1.Visible = False
    Command2.Visible = False
    Label2.Caption = "Gloss�rio"
    Label1(0).Caption = "  X-Zombies-X = Impede os Zombies de atacarem."
    Label1(1).Caption = "  X-Ataque-X = Criaturas n�o atacam neste turno"
    Label1(2).Caption = "  Inbloque�vel = Criaturas com esta abilidade n�o podem ser bloqueadas"
    Label1(3).Caption = "  Abilidade de Voar = S� podem ser bloqueadas por criaturas com a mesma abilidade"
    Label1(4).Caption = "  X-Magia Inimiga de Ataque-X = Bloqueia uma magia inimiga causadora de dano"
    Label1(5).Caption = "  X-Ac. +Ataque/+Vida-X = As criaturas com poder inferior a X/X n�o podem atacar"
    Label1(6).Caption = "  +1 Vida - Jogador/Vida Toda - Criatura = +1 vida ao jogador ou regenera uma criatura"
    Label1(7).Caption = "  X-Magia Inimiga-X = Bloqueia qualquer magia inimiga"
    Label1(8).Caption = "  <=Reflectir Magias Inimigas<= Os feiti�os viram-se contra o feiticeiro"
    Label1(9).Caption = "  X-Veneno-X = O jogador ou criatura alvo s�o desentoxicados"
End Sub

Private Sub h_Click()
    Form8.Visible = False
End Sub

